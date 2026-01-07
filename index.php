<?php
declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');
header('Cache-Control: no-store');

function send_json(int $status, array $payload): void
{
    http_response_code($status);
    echo json_encode($payload, JSON_UNESCAPED_UNICODE);
    exit;
}

function get_request_path(): string
{
    if (!empty($_SERVER['PATH_INFO'])) {
        return $_SERVER['PATH_INFO'];
    }

    $uri = $_SERVER['REQUEST_URI'] ?? '/';
    $script = $_SERVER['SCRIPT_NAME'] ?? '';
    if ($script !== '' && strpos($uri, $script) === 0) {
        $path = substr($uri, strlen($script));
    } else {
        $base = rtrim(dirname($script), '/');
        if ($base !== '' && strpos($uri, $base) === 0) {
            $path = substr($uri, strlen($base));
        } else {
            $path = $uri;
        }
    }

    $path = strtok($path, '?');
    if ($path === false || $path === '') {
        return '/';
    }

    return $path[0] === '/' ? $path : '/' . $path;
}

function get_pdo(): PDO
{
    $envPath = __DIR__ . '/.env.local';
    $env = [];
    if (is_file($envPath)) {
        $env = parse_ini_file($envPath, false, INI_SCANNER_RAW) ?: [];
    }

    $host = $env['DB_HOST'] ?? 'localhost';
    $db = $env['DB_NAME'] ?? 'slms';
    $user = $env['DB_USER'] ?? 'root';
    $pass = $env['DB_PASS'] ?? '';
    $dsn = "mysql:host={$host};dbname={$db};charset=utf8mb4";

    return new PDO($dsn, $user, $pass, [
        PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
        PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
    ]);
}

function read_json_body(): array
{
    $raw = file_get_contents('php://input');
    if ($raw === false || trim($raw) === '') {
        return [];
    }
    $data = json_decode($raw, true);
    return is_array($data) ? $data : [];
}

function normalize_header(string $value): string
{
    $value = trim($value);
    $value = preg_replace('/\s+/u', '', $value);
    $value = preg_replace('/[^\p{L}\p{N}]+/u', '', $value);
    return mb_strtolower($value, 'UTF-8');
}

function excel_serial_to_date(int $serial): string
{
    $base = new DateTime('1899-12-30');
    $base->modify('+' . $serial . ' days');
    return $base->format('Y-m-d');
}

function adjust_thai_year(string $dateStr): string
{
    $dt = DateTime::createFromFormat('Y-m-d', $dateStr);
    if (!$dt instanceof DateTime) {
        return $dateStr;
    }

    $year = (int) $dt->format('Y');
    if ($year >= 2400) {
        $dt->modify('-543 years');
    }

    return $dt->format('Y-m-d');
}

function parse_date_value($value): ?string
{
    if ($value === null || $value === '') {
        return null;
    }

    if (is_numeric($value)) {
        return adjust_thai_year(excel_serial_to_date((int) $value));
    }

    $value = trim((string) $value);
    $value = str_replace(['.', '\\'], ['/', '/'], $value);
    $value = preg_replace('/\s+/', ' ', $value);

    $dt = DateTime::createFromFormat('d/m/Y', $value);
    if ($dt instanceof DateTime) {
        return adjust_thai_year($dt->format('Y-m-d'));
    }

    $dt = DateTime::createFromFormat('d-m-Y', $value);
    if ($dt instanceof DateTime) {
        return adjust_thai_year($dt->format('Y-m-d'));
    }

    $timestamp = strtotime($value);
    if ($timestamp !== false) {
        return adjust_thai_year(date('Y-m-d', $timestamp));
    }

    return null;
}

function column_to_index(string $cellRef): int
{
    $letters = preg_replace('/[^A-Z]/', '', strtoupper($cellRef));
    $index = 0;
    for ($i = 0, $len = strlen($letters); $i < $len; $i++) {
        $index = $index * 26 + (ord($letters[$i]) - 64);
    }
    return $index - 1;
}

function read_xlsx_rows(string $filePath): array
{
    if (!class_exists('ZipArchive')) {
        throw new RuntimeException('ZIP extension not available');
    }

    $zip = new ZipArchive();
    if ($zip->open($filePath) !== true) {
        throw new RuntimeException('Unable to open xlsx file');
    }

    $sharedStrings = [];
    $sharedXml = $zip->getFromName('xl/sharedStrings.xml');
    if ($sharedXml !== false) {
        libxml_use_internal_errors(true);
        $shared = simplexml_load_string($sharedXml);
        libxml_clear_errors();
        if ($shared && isset($shared->si)) {
            foreach ($shared->si as $si) {
                if (isset($si->t)) {
                    $sharedStrings[] = (string) $si->t;
                } elseif (isset($si->r)) {
                    $text = '';
                    foreach ($si->r as $run) {
                        $text .= (string) $run->t;
                    }
                    $sharedStrings[] = $text;
                }
            }
        }
    }

    $sheetPath = 'xl/worksheets/sheet1.xml';
    $workbookXml = $zip->getFromName('xl/workbook.xml');
    if ($workbookXml !== false) {
        libxml_use_internal_errors(true);
        $workbook = simplexml_load_string($workbookXml);
        libxml_clear_errors();
        if ($workbook && isset($workbook->sheets->sheet)) {
            $firstSheet = $workbook->sheets->sheet[0];
            $rid = (string) $firstSheet->attributes('r', true)->id;
            if ($rid !== '') {
                $relsXml = $zip->getFromName('xl/_rels/workbook.xml.rels');
                if ($relsXml !== false) {
                    libxml_use_internal_errors(true);
                    $rels = simplexml_load_string($relsXml);
                    libxml_clear_errors();
                    if ($rels && isset($rels->Relationship)) {
                        foreach ($rels->Relationship as $rel) {
                            if ((string) $rel['Id'] === $rid) {
                                $target = (string) $rel['Target'];
                                if ($target !== '') {
                                    $sheetPath = strpos($target, 'xl/') === 0 ? $target : 'xl/' . ltrim($target, '/');
                                }
                                break;
                            }
                        }
                    }
                }
            }
        }
    }

    $sheetXml = $zip->getFromName($sheetPath);
    if ($sheetXml === false) {
        throw new RuntimeException('Worksheet not found');
    }

    $sheet = simplexml_load_string($sheetXml);
    if (!$sheet || !isset($sheet->sheetData->row)) {
        throw new RuntimeException('Invalid worksheet');
    }

    $rows = [];
    foreach ($sheet->sheetData->row as $row) {
        $cells = [];
        foreach ($row->c as $cell) {
            $ref = (string) $cell['r'];
            $index = column_to_index($ref);
            $type = (string) $cell['t'];
            $value = isset($cell->v) ? (string) $cell->v : '';

            if ($type === 's') {
                $value = $sharedStrings[(int) $value] ?? '';
            } elseif ($type === 'inlineStr' && isset($cell->is->t)) {
                $value = (string) $cell->is->t;
            }

            $cells[$index] = $value;
        }

        if ($cells) {
            ksort($cells);
            $maxIndex = array_key_last($cells);
            $row = [];
            for ($i = 0; $i <= $maxIndex; $i++) {
                $row[] = $cells[$i] ?? '';
            }
            $rows[] = $row;
        }
    }

    $zip->close();
    return $rows;
}

function build_header_map_from_rows(array $rows, int $maxRows = 10): array
{
    $map = [];
    $headerMapRaw = [
        'cid' => 'cid',
        'full_name' => 'full_name',
        'position_level' => 'position_level',
        'position_no' => 'position_no',
        'workplace' => 'workplace',
        'program' => 'program',
        'program_years' => 'program_years',
        'institute' => 'institute',
        'start_date' => 'start_date',
        'end_date' => 'end_date',
        'note' => 'note',
        'order_no' => 'order_no',
        'ชื่อสกุล' => 'full_name',
        'ชื่อ-สกุล' => 'full_name',
        'ตำแหน่งส่วนราชการตามว18' => 'position_level',
        'ตำแหน่งส่วนราชการตามว๑๘' => 'position_level',
        'ตำแหน่งส่วนราชการตามจ18' => 'position_level',
        'ตำแหน่งส่วนราชการตามจ๑๘' => 'position_level',
        'ตำแหน่งเลขที่' => 'position_no',
        'สถานที่ปฏิบัติงานจริง' => 'workplace',
        'หลักสูตร' => 'program',
        'หลักสูตรปี' => 'program_years',
        'หลักสูตร(ปี)' => 'program_years',
        'สถานที่ศึกษา' => 'institute',
        'เริ่มต้นวด้ป' => 'start_date',
        'ตั้งแต่วดป' => 'start_date',
        'สิ้นสุดวด้ป' => 'end_date',
        'ถึงวดป' => 'end_date',
        'หมายเหตุ' => 'note',
        'โควตา' => 'note',
        'เลขที่คำสั่ง' => 'order_no',
    ];

    $headerMap = [];
    foreach ($headerMapRaw as $label => $field) {
        $normalized = normalize_header((string) $label);
        if (!isset($headerMap[$normalized])) {
            $headerMap[$normalized] = $field;
        }
    }

    $limit = min($maxRows, count($rows));
    for ($rowIndex = 0; $rowIndex < $limit; $rowIndex++) {
        $row = $rows[$rowIndex];
        foreach ($row as $index => $label) {
            if ($label === null || $label === '') {
                continue;
            }
            $normalized = normalize_header((string) $label);
            if (isset($headerMap[$normalized])) {
                $field = $headerMap[$normalized];
                if (!isset($map[$field])) {
                    $map[$field] = $index;
                }
            }
        }
    }

    return $map;
}

function find_data_start(array $rows, array $map): int
{
    $limit = min(20, count($rows));
    for ($i = 0; $i < $limit; $i++) {
        $fullName = get_cell($rows[$i], $map, 'full_name');
        $position = get_cell($rows[$i], $map, 'position_level');
        if ($fullName !== null && trim($fullName) !== '' && $position !== null && trim($position) !== '') {
            return $i;
        }
    }

    return -1;
}

function get_cell(array $row, array $map, string $key): ?string
{
    if (!isset($map[$key])) {
        return null;
    }
    $index = $map[$key];
    return isset($row[$index]) ? trim((string) $row[$index]) : null;
}

function compute_leave_status(string $startDate, string $endDate): string
{
    $today = new DateTime('today');
    $start = new DateTime($startDate);
    $end = new DateTime($endDate);

    if ($today < $start) {
        return 'pending';
    }
    if ($today > $end) {
        return 'completed';
    }
    return 'active';
}

function extract_position_title(string $value): string
{
    $raw = trim($value);
    if ($raw === '') {
        return 'ไม่ระบุตำแหน่ง';
    }
    $parts = preg_split("/\r?\n+/", $raw);
    if (!$parts) {
        $normalized = preg_replace('/\s+/', ' ', $raw);
        return $normalized ?? $raw;
    }
    $title = trim($parts[0]) !== '' ? trim($parts[0]) : $raw;
    $normalized = preg_replace('/\s+/', ' ', $title);
    return $normalized ?? $title;
}

function map_leave_row(array $row): array
{
    $status = compute_leave_status($row['start_date'], $row['end_date']);

    return [
        'id' => (int) $row['id'],
        'cid' => $row['cid'],
        'position_level' => $row['position_level'],
        'position_no' => $row['position_no'],
        'workplace' => $row['workplace'],
        'program' => $row['program'],
        'program_years' => (int) $row['program_years'],
        'institute' => $row['institute'],
        'note' => $row['note'],
        'full_name' => $row['full_name'],
        'position' => $row['position_level'],
        'order_no' => $row['order_no'],
        'level' => $row['program'],
        'type' => $row['program_years'] . ' ปี',
        'start_date' => $row['start_date'],
        'end_date' => $row['end_date'],
        'status' => $status,
    ];
}

$path = get_request_path();
$method = $_SERVER['REQUEST_METHOD'] ?? 'GET';

if ($path === '/' || $path === '/api') {
    send_json(200, ['status' => 'ok']);
}

switch ($path) {
    case '/api/dashboard':
        try {
            $pdo = get_pdo();
            $rows = $pdo->query('SELECT position_level, start_date, end_date FROM study_leaves')->fetchAll();
            $imports = $pdo->query('SELECT original_name, stored_path, inserted, skipped, created_at FROM import_logs ORDER BY created_at DESC LIMIT 5')->fetchAll();
        } catch (Throwable $e) {
            send_json(500, ['error' => 'Database error']);
        }

        $statusFilter = $_GET['status'] ?? 'all';
        $allowed = ['all', 'active', 'pending', 'completed'];
        if (!in_array($statusFilter, $allowed, true)) {
            $statusFilter = 'all';
        }

        $now = new DateTime('today');
        $dueLimit = (clone $now)->modify('+90 days');
        $total = 0;
        $due = 0;
        $positionCounts = [];

        foreach ($rows as $row) {
            $status = compute_leave_status($row['start_date'], $row['end_date']);
            if ($statusFilter !== 'all' && $status !== $statusFilter) {
                continue;
            }

            $total++;
            $endDate = new DateTime($row['end_date']);
            if ($endDate >= $now && $endDate <= $dueLimit) {
                $due++;
            }

            $position = extract_position_title((string) ($row['position_level'] ?? ''));
            if (!isset($positionCounts[$position])) {
                $positionCounts[$position] = 0;
            }
            $positionCounts[$position]++;
        }

        arsort($positionCounts);
        $topPositions = [];
        foreach ($positionCounts as $position => $count) {
            $topPositions[] = ['position' => $position, 'count' => $count];
            if (count($topPositions) >= 4) {
                break;
            }
        }

        send_json(200, [
            'data' => [
                'total_leaves' => $total,
                'suspension_active' => 0,
                'due_to_reinstate' => $due,
                'suspension_amount' => 0,
                'top_positions' => $topPositions,
                'recent_imports' => $imports,
            ],
        ]);
        break;

    case '/api/leaves':
        if ($method === 'POST') {
            $payload = read_json_body();
            $isUpdate = !empty($payload['id']);
            $requiredFields = [
                'cid',
                'full_name',
                'position_level',
                'position_no',
                'workplace',
                'program',
                'program_years',
                'institute',
                'start_date',
                'end_date',
                'order_no',
            ];
            foreach ($requiredFields as $field) {
                if (empty($payload[$field])) {
                    send_json(400, ['error' => 'Missing field: ' . $field]);
                }
            }

            $startDate = parse_date_value($payload['start_date']);
            $endDate = parse_date_value($payload['end_date']);
            if ($startDate === null || $endDate === null) {
                send_json(400, ['error' => 'Invalid date format']);
            }

            $programYears = (int) preg_replace('/[^0-9]/', '', (string) $payload['program_years']);
            if ($programYears <= 0) {
                $programYears = 1;
            }

            try {
                $pdo = get_pdo();
                if ($isUpdate) {
                    $stmt = $pdo->prepare(
                        'UPDATE study_leaves
                         SET cid = :cid,
                             full_name = :full_name,
                             position_level = :position_level,
                             position_no = :position_no,
                             workplace = :workplace,
                             program = :program,
                             program_years = :program_years,
                             institute = :institute,
                             start_date = :start_date,
                             end_date = :end_date,
                             note = :note,
                             order_no = :order_no
                         WHERE id = :id'
                    );
                    $stmt->execute([
                        ':id' => (int) $payload['id'],
                        ':cid' => $payload['cid'],
                        ':full_name' => $payload['full_name'],
                        ':position_level' => $payload['position_level'],
                        ':position_no' => $payload['position_no'],
                        ':workplace' => $payload['workplace'],
                        ':program' => $payload['program'],
                        ':program_years' => $programYears,
                        ':institute' => $payload['institute'],
                        ':start_date' => $startDate,
                        ':end_date' => $endDate,
                        ':note' => $payload['note'] ?? null,
                        ':order_no' => $payload['order_no'],
                    ]);
                } else {
                    $stmt = $pdo->prepare(
                        'INSERT INTO study_leaves (cid, full_name, position_level, position_no, workplace, program, program_years, institute, start_date, end_date, note, order_no)
                         VALUES (:cid, :full_name, :position_level, :position_no, :workplace, :program, :program_years, :institute, :start_date, :end_date, :note, :order_no)'
                    );
                    $stmt->execute([
                        ':cid' => $payload['cid'],
                        ':full_name' => $payload['full_name'],
                        ':position_level' => $payload['position_level'],
                        ':position_no' => $payload['position_no'],
                        ':workplace' => $payload['workplace'],
                        ':program' => $payload['program'],
                        ':program_years' => $programYears,
                        ':institute' => $payload['institute'],
                        ':start_date' => $startDate,
                        ':end_date' => $endDate,
                        ':note' => $payload['note'] ?? null,
                        ':order_no' => $payload['order_no'],
                    ]);
                }
            } catch (Throwable $e) {
                send_json(500, ['error' => 'Database error']);
            }

            send_json(200, ['success' => true]);
        }

        try {
            $pdo = get_pdo();
            $stmt = $pdo->query('SELECT * FROM study_leaves ORDER BY id DESC');
            $rows = $stmt->fetchAll();
        } catch (Throwable $e) {
            send_json(500, ['error' => 'Database error']);
        }

        $data = array_map('map_leave_row', $rows);
        send_json(200, ['data' => $data]);
        break;

    case '/api/suspensions':
        send_json(200, ['data' => []]);
        break;

    case '/api/reports/summary':
        try {
            $pdo = get_pdo();
            $rows = $pdo->query('SELECT start_date, end_date, program_years FROM study_leaves')->fetchAll();
        } catch (Throwable $e) {
            send_json(500, ['error' => 'Database error']);
        }

        $total = count($rows);
        $full = 0;
        $part = 0;
        $due = 0;
        $statusCounts = [
            'active' => 0,
            'pending' => 0,
            'completed' => 0,
        ];
        $now = new DateTime('today');
        $dueLimit = (clone $now)->modify('+90 days');

        foreach ($rows as $row) {
            $programYears = (int) ($row['program_years'] ?? 0);
            if ($programYears >= 2) {
                $full++;
            } else {
                $part++;
            }

            $status = compute_leave_status($row['start_date'], $row['end_date']);
            if (isset($statusCounts[$status])) {
                $statusCounts[$status]++;
            }

            $endDate = new DateTime($row['end_date']);
            if ($endDate >= $now && $endDate <= $dueLimit) {
                $due++;
            }
        }

        send_json(200, [
            'data' => [
                'budget_savings' => ['total' => 0],
                'leave_counts' => [
                    'total' => $total,
                    'full_time' => $full,
                    'part_time' => $part,
                ],
                'status_counts' => $statusCounts,
                'due_reinstates' => $due,
                'dept_stats' => [],
            ],
        ]);
        break;

    case '/api/import/upload':
        if ($method !== 'POST') {
            send_json(405, ['error' => 'Method not allowed']);
        }
        if (empty($_FILES['file'])) {
            send_json(400, ['error' => 'No file uploaded']);
        }

        $uploadDir = __DIR__ . '/public/uploads/';
        if (!is_dir($uploadDir) && !@mkdir($uploadDir, 0755, true)) {
            send_json(500, ['error' => 'Failed to create upload directory']);
        }

        $file = $_FILES['file'];
        $ext = strtolower(pathinfo($file['name'] ?? '', PATHINFO_EXTENSION));
        if ($ext !== 'pdf') {
            send_json(400, ['error' => 'Only PDF is allowed']);
        }

        $name = bin2hex(random_bytes(16)) . '.pdf';
        $target = $uploadDir . $name;
        if (!move_uploaded_file($file['tmp_name'], $target)) {
            send_json(500, ['error' => 'Failed to move uploaded file']);
        }

        send_json(200, [
            'success' => true,
            'originalName' => $file['name'],
            'path' => 'public/uploads/' . $name,
            'document_id' => uniqid('doc_', true),
        ]);
        break;

    case '/api/import/excel':
        if ($method !== 'POST') {
            send_json(405, ['error' => 'Method not allowed']);
        }
        if (empty($_FILES['file'])) {
            send_json(400, ['error' => 'No file uploaded']);
        }

        $file = $_FILES['file'];
        $ext = strtolower(pathinfo($file['name'] ?? '', PATHINFO_EXTENSION));
        if ($ext !== 'xlsx') {
            send_json(400, ['error' => 'Only .xlsx is allowed']);
        }

        $uploadDir = __DIR__ . '/public/uploads/';
        if (!is_dir($uploadDir) && !@mkdir($uploadDir, 0755, true)) {
            send_json(500, ['error' => 'Failed to create upload directory']);
        }

        $name = bin2hex(random_bytes(16)) . '.xlsx';
        $target = $uploadDir . $name;
        if (!move_uploaded_file($file['tmp_name'], $target)) {
            send_json(500, ['error' => 'Failed to move uploaded file']);
        }

        try {
            $rows = read_xlsx_rows($target);
        } catch (Throwable $e) {
            send_json(400, ['error' => 'Unable to read Excel file: ' . $e->getMessage()]);
        }

        if (!$rows) {
            send_json(400, ['error' => 'Excel file is empty']);
        }

        $required = ['cid', 'full_name', 'position_level', 'position_no', 'workplace', 'program', 'program_years', 'institute', 'start_date', 'end_date', 'order_no'];
        $headerMap = build_header_map_from_rows($rows, 10);
        $missing = array_diff($required, array_keys($headerMap));
        if ($missing) {
            send_json(400, [
                'error' => 'Missing required columns',
                'missing' => array_values($missing),
                'expected' => [
                    'cid',
                    'ชื่อ-สกุล',
                    'ตำแหน่ง/ส่วนราชการตาม จ.18',
                    'ตำแหน่งเลขที่',
                    'สถานที่ปฏิบัติงานจริง',
                    'หลักสูตร',
                    'หลักสูตร(ปี)',
                    'สถานที่ศึกษา',
                    'ตั้งแต่ (ว.ด.ป.)',
                    'ถึง (ว.ด.ป.)',
                    'หมายเหตุ',
                    'เลขที่คำสั่ง',
                ],
            ]);
        }

        $dataStart = find_data_start($rows, $headerMap);
        if ($dataStart < 0) {
            send_json(400, ['error' => 'Unable to find data rows']);
        }

        try {
            $pdo = get_pdo();
            $pdo->beginTransaction();
            $stmt = $pdo->prepare(
                'INSERT INTO study_leaves (cid, full_name, position_level, position_no, workplace, program, program_years, institute, start_date, end_date, note, order_no)
                 VALUES (:cid, :full_name, :position_level, :position_no, :workplace, :program, :program_years, :institute, :start_date, :end_date, :note, :order_no)'
            );

            $inserted = 0;
            $skipped = 0;

            foreach (array_slice($rows, $dataStart) as $row) {
                $cid = get_cell($row, $headerMap, 'cid');
                $fullName = get_cell($row, $headerMap, 'full_name');
                if ($cid === null && $fullName === null) {
                    $skipped++;
                    continue;
                }

                $startDate = parse_date_value(get_cell($row, $headerMap, 'start_date'));
                $endDate = parse_date_value(get_cell($row, $headerMap, 'end_date'));
                if ($startDate === null || $endDate === null) {
                    $skipped++;
                    continue;
                }

                $programYearsRaw = get_cell($row, $headerMap, 'program_years');
                $programYears = (int) preg_replace('/[^0-9]/', '', (string) $programYearsRaw);
                if ($programYears <= 0) {
                    $programYears = 1;
                }

                $stmt->execute([
                    ':cid' => $cid ?? '',
                    ':full_name' => $fullName ?? '',
                    ':position_level' => get_cell($row, $headerMap, 'position_level') ?? '',
                    ':position_no' => get_cell($row, $headerMap, 'position_no') ?? '',
                    ':workplace' => get_cell($row, $headerMap, 'workplace') ?? '',
                    ':program' => get_cell($row, $headerMap, 'program') ?? '',
                    ':program_years' => $programYears,
                    ':institute' => get_cell($row, $headerMap, 'institute') ?? '',
                    ':start_date' => $startDate,
                    ':end_date' => $endDate,
                    ':note' => get_cell($row, $headerMap, 'note'),
                    ':order_no' => get_cell($row, $headerMap, 'order_no') ?? '',
                ]);
                $inserted++;
            }

            $pdo->commit();
        } catch (Throwable $e) {
            if (isset($pdo) && $pdo->inTransaction()) {
                $pdo->rollBack();
            }
            send_json(500, ['error' => 'Database error']);
        }

        try {
            $pdo = get_pdo();
            $logStmt = $pdo->prepare(
                'INSERT INTO import_logs (original_name, stored_path, inserted, skipped)
                 VALUES (:original_name, :stored_path, :inserted, :skipped)'
            );
            $logStmt->execute([
                ':original_name' => $file['name'],
                ':stored_path' => 'public/uploads/' . $name,
                ':inserted' => $inserted,
                ':skipped' => $skipped,
            ]);
        } catch (Throwable $e) {
            // Log failure should not block import response.
        }

        send_json(200, [
            'success' => true,
            'originalName' => $file['name'],
            'path' => 'public/uploads/' . $name,
            'inserted' => $inserted,
            'skipped' => $skipped,
        ]);
        break;
}

send_json(404, ['error' => 'Not found']);
