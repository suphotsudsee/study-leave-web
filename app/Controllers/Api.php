<?php

namespace App\Controllers;

use CodeIgniter\HTTP\ResponseInterface;
use Config\Database;
use DateTime;
use RuntimeException;
use Throwable;
use ZipArchive;

class Api extends BaseController
{
    private function json(array $payload, int $status = 200): ResponseInterface
    {
        return $this->response->setStatusCode($status)->setJSON($payload);
    }

    private function readJson(): array
    {
        $data = $this->request->getJSON(true);
        return is_array($data) ? $data : [];
    }

    private function normalizeHeader(string $value): string
    {
        $value = trim($value);
        $value = preg_replace('/\s+/u', '', $value);
        $value = preg_replace('/[^\p{L}\p{N}]+/u', '', $value);
        return mb_strtolower($value, 'UTF-8');
    }

    private function excelSerialToDate(int $serial): string
    {
        $base = new DateTime('1899-12-30');
        $base->modify('+' . $serial . ' days');
        return $base->format('Y-m-d');
    }

    private function adjustThaiYear(string $dateStr): string
    {
        $dt = DateTime::createFromFormat('Y-m-d', $dateStr);
        if (! $dt instanceof DateTime) {
            return $dateStr;
        }

        $year = (int) $dt->format('Y');
        if ($year >= 2400) {
            $dt->modify('-543 years');
        }

        return $dt->format('Y-m-d');
    }

    private function replaceThaiMonths(string $value): string
    {
        $monthMap = [
            'ม\\.ค\\.?|มกราคม' => '1',
            'ก\\.พ\\.?|กุมภาพันธ์' => '2',
            'มี\\.ค\\.?|มีนาคม' => '3',
            'เม\\.ย\\.?|เมษายน' => '4',
            'พ\\.ค\\.?|พฤษภาคม' => '5',
            'มิ\\.ย\\.?|มิถุนายน' => '6',
            'ก\\.ค\\.?|กรกฎาคม' => '7',
            'ส\\.ค\\.?|สิงหาคม' => '8',
            'ก\\.ย\\.?|กันยายน' => '9',
            'ต\\.ค\\.?|ตุลาคม' => '10',
            'พ\\.ย\\.?|พฤศจิกายน' => '11',
            'ธ\\.ค\\.?|ธันวาคม' => '12',
        ];

        foreach ($monthMap as $pattern => $month) {
            $value = preg_replace('/' . $pattern . '/u', $month, $value);
        }

        return $value;
    }

    private function looksLikeDateCandidate(string $value): bool
    {
        $trimmed = trim($value);
        if ($trimmed === '') {
            return false;
        }

        if (is_numeric($trimmed)) {
            return (int) $trimmed >= 20000;
        }

        return preg_match('/[\\/-]|[ก-ฮ]/u', $trimmed) === 1;
    }

    private function parseDateValue($value): ?string
    {
        if ($value === null || $value === '') {
            return null;
        }

        if (is_numeric($value)) {
            return $this->adjustThaiYear($this->excelSerialToDate((int) $value));
        }

        $value = trim((string) $value);
        $value = strtr($value, [
            '๐' => '0',
            '๑' => '1',
            '๒' => '2',
            '๓' => '3',
            '๔' => '4',
            '๕' => '5',
            '๖' => '6',
            '๗' => '7',
            '๘' => '8',
            '๙' => '9',
        ]);
        $value = str_replace(['.', '\\'], ['/', '/'], $value);
        $value = preg_replace('/\s+/', ' ', $value);
        $value = $this->replaceThaiMonths($value);
        $value = preg_replace('/\s+/', '/', $value);

        $dt = DateTime::createFromFormat('d/m/Y', $value);
        if ($dt instanceof DateTime) {
            return $this->adjustThaiYear($dt->format('Y-m-d'));
        }

        $dt = DateTime::createFromFormat('d-m-Y', $value);
        if ($dt instanceof DateTime) {
            return $this->adjustThaiYear($dt->format('Y-m-d'));
        }

        $timestamp = strtotime($value);
        if ($timestamp !== false) {
            return $this->adjustThaiYear(date('Y-m-d', $timestamp));
        }

        return null;
    }

    private function columnToIndex(string $cellRef): int
    {
        $letters = preg_replace('/[^A-Z]/', '', strtoupper($cellRef));
        $index = 0;
        for ($i = 0, $len = strlen($letters); $i < $len; $i++) {
            $index = $index * 26 + (ord($letters[$i]) - 64);
        }
        return $index - 1;
    }

    private function readXlsxRows(string $filePath): array
    {
        if (! class_exists(ZipArchive::class)) {
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
        if (! $sheet || ! isset($sheet->sheetData->row)) {
            throw new RuntimeException('Invalid worksheet');
        }

        $rows = [];
        foreach ($sheet->sheetData->row as $row) {
            $cells = [];
            foreach ($row->c as $cell) {
                $ref = (string) $cell['r'];
                $index = $this->columnToIndex($ref);
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
                $rowValues = [];
                for ($i = 0; $i <= $maxIndex; $i++) {
                    $rowValues[] = $cells[$i] ?? '';
                }
                $rows[] = $rowValues;
            }
        }

        $zip->close();
        return $rows;
    }

    private function buildHeaderMapFromRows(array $rows, int $maxRows = 10): array
    {
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
            'ปีที่ขออนุมัติ' => 'approval_year',
            'หมายเหตุ' => 'note',
            'โควตา' => 'note',
            'เลขที่คำสั่ง' => 'order_no',
        ];

        $headerMap = [];
        foreach ($headerMapRaw as $label => $field) {
            $normalized = $this->normalizeHeader((string) $label);
            if (! isset($headerMap[$normalized])) {
                $headerMap[$normalized] = $field;
            }
        }

        $map = [];
        $limit = min($maxRows, count($rows));
        for ($rowIndex = 0; $rowIndex < $limit; $rowIndex++) {
            $row = $rows[$rowIndex];
            foreach ($row as $index => $label) {
                if ($label === null || $label === '') {
                    continue;
                }
                $normalized = $this->normalizeHeader((string) $label);
                if (isset($headerMap[$normalized])) {
                    $field = $headerMap[$normalized];
                    if (! isset($map[$field])) {
                        $map[$field] = $index;
                    }
                }
            }
        }

        return $map;
    }

    private function getCell(array $row, array $map, string $key): ?string
    {
        if (! isset($map[$key])) {
            return null;
        }
        $index = $map[$key];
        return isset($row[$index]) ? trim((string) $row[$index]) : null;
    }

    private function findDataStart(array $rows, array $map): int
    {
        $limit = min(20, count($rows));
        for ($i = 0; $i < $limit; $i++) {
            $fullName = $this->getCell($rows[$i], $map, 'full_name');
            $position = $this->getCell($rows[$i], $map, 'position_level');
            if ($fullName !== null && trim($fullName) !== '' && $position !== null && trim($position) !== '') {
                return $i;
            }
        }

        return -1;
    }

    private function computeLeaveStatus(string $startDate, string $endDate): string
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

    private function extractPositionTitle(string $value): string
    {
        $raw = trim($value);
        if ($raw === '') {
            return 'ไม่ระบุตำแหน่ง';
        }
        $parts = preg_split("/\r?\n+/", $raw);
        if (! $parts) {
            $normalized = preg_replace('/\s+/', ' ', $raw);
            return $normalized ?? $raw;
        }
        $title = trim($parts[0]) !== '' ? trim($parts[0]) : $raw;
        $normalized = preg_replace('/\s+/', ' ', $title);
        return $normalized ?? $title;
    }

    public function dashboard(): ResponseInterface
    {
        try {
            $db = Database::connect();
            $rows = $db->query('SELECT position_level, start_date, end_date FROM study_leaves')->getResultArray();
            $imports = $db->query('SELECT original_name, stored_path, inserted, skipped, created_at FROM import_logs ORDER BY created_at DESC LIMIT 5')->getResultArray();
        } catch (Throwable $e) {
            return $this->json(['error' => 'Database error'], 500);
        }

        $statusFilter = $this->request->getGet('status') ?? 'all';
        $allowed = ['all', 'active', 'pending', 'completed'];
        if (! in_array($statusFilter, $allowed, true)) {
            $statusFilter = 'all';
        }

        $now = new DateTime('today');
        $dueLimit = (clone $now)->modify('+90 days');
        $total = 0;
        $due = 0;
        $positionCounts = [];

        foreach ($rows as $row) {
            $status = $this->computeLeaveStatus($row['start_date'], $row['end_date']);
            if ($statusFilter !== 'all' && $status !== $statusFilter) {
                continue;
            }
            $total++;
            $endDate = new DateTime($row['end_date']);
            if ($endDate >= $now && $endDate <= $dueLimit) {
                $due++;
            }

            $position = $this->extractPositionTitle((string) ($row['position_level'] ?? ''));
            if (! isset($positionCounts[$position])) {
                $positionCounts[$position] = 0;
            }
            $positionCounts[$position]++;
        }

        arsort($positionCounts);
        $topPositions = [];
        $topTotal = 0;
        foreach ($positionCounts as $position => $count) {
            $topPositions[] = ['position' => $position, 'count' => $count];
            $topTotal += $count;
            if (count($topPositions) >= 4) {
                break;
            }
        }
        $otherPositions = max(0, $total - $topTotal);

        return $this->json([
            'data' => [
                'total_leaves' => $total,
                'suspension_active' => 0,
                'due_to_reinstate' => $due,
                'suspension_amount' => 0,
                'top_positions' => $topPositions,
                'other_positions' => $otherPositions,
                'recent_imports' => $imports,
            ],
        ]);
    }

    public function leaves(): ResponseInterface
    {
        $method = strtoupper($this->request->getMethod());
        $db = Database::connect();

        if ($method === 'POST') {
            $payload = $this->readJson();
            $isUpdate = ! empty($payload['id']);
            $required = [
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
            foreach ($required as $field) {
                if (empty($payload[$field])) {
                    return $this->json(['error' => 'Missing field: ' . $field], 400);
                }
            }

            $startDate = $this->parseDateValue($payload['start_date']);
            $endDate = $this->parseDateValue($payload['end_date']);
            if ($startDate === null || $endDate === null) {
                return $this->json(['error' => 'Invalid date format'], 400);
            }

            $programYearsRaw = $payload['program_years'] ?? '';
            $programYears = (int) preg_replace('/[^0-9]/', '', (string) $programYearsRaw);
            if ($programYears <= 0) {
                $programYears = 1;
            }

            $data = [
                'cid' => $payload['cid'],
                'full_name' => $payload['full_name'],
                'position_level' => $payload['position_level'],
                'position_no' => $payload['position_no'],
                'workplace' => $payload['workplace'],
                'program' => $payload['program'],
                'program_years' => $programYears,
                'institute' => $payload['institute'],
                'start_date' => $startDate,
                'end_date' => $endDate,
                'note' => $payload['note'] ?? null,
                'order_no' => $payload['order_no'],
            ];

            try {
                $builder = $db->table('study_leaves');
                if ($isUpdate) {
                    $builder->where('id', (int) $payload['id'])->update($data);
                } else {
                    $builder->insert($data);
                }
            } catch (Throwable $e) {
                return $this->json(['error' => 'Database error'], 500);
            }

            return $this->json(['success' => true]);
        }

        try {
            $rows = $db->query('SELECT * FROM study_leaves ORDER BY id DESC')->getResultArray();
        } catch (Throwable $e) {
            return $this->json(['error' => 'Database error'], 500);
        }

        $data = array_map(function (array $row): array {
            $status = $this->computeLeaveStatus($row['start_date'], $row['end_date']);
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
        }, $rows);

        return $this->json(['data' => $data]);
    }

    public function reportsSummary(): ResponseInterface
    {
        try {
            $db = Database::connect();
            $rows = $db->query('SELECT start_date, end_date, program_years FROM study_leaves')->getResultArray();
        } catch (Throwable $e) {
            return $this->json(['error' => 'Database error'], 500);
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

            $status = $this->computeLeaveStatus($row['start_date'], $row['end_date']);
            if (isset($statusCounts[$status])) {
                $statusCounts[$status]++;
            }

            $endDate = new DateTime($row['end_date']);
            if ($endDate >= $now && $endDate <= $dueLimit) {
                $due++;
            }
        }

        return $this->json([
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
    }

    public function importExcel(): ResponseInterface
    {
        $file = $this->request->getFile('file');
        if (! $file || ! $file->isValid()) {
            return $this->json(['error' => 'No file uploaded'], 400);
        }

        if (strtolower($file->getClientExtension()) !== 'xlsx') {
            return $this->json(['error' => 'Only .xlsx is allowed'], 400);
        }

        $uploadDir = FCPATH . 'uploads';
        if (! is_dir($uploadDir) && ! @mkdir($uploadDir, 0755, true)) {
            return $this->json(['error' => 'Failed to create upload directory'], 500);
        }

        $newName = bin2hex(random_bytes(16)) . '.xlsx';
        $file->move($uploadDir, $newName);
        $storedPath = 'public/uploads/' . $newName;

        try {
            $rows = $this->readXlsxRows($uploadDir . DIRECTORY_SEPARATOR . $newName);
        } catch (Throwable $e) {
            return $this->json(['error' => 'Unable to read Excel file: ' . $e->getMessage()], 400);
        }

        if (! $rows) {
            return $this->json(['error' => 'Excel file is empty'], 400);
        }

        $required = ['cid', 'full_name', 'position_level', 'position_no', 'workplace', 'program', 'program_years', 'institute', 'start_date', 'end_date', 'order_no'];
        $headerMap = $this->buildHeaderMapFromRows($rows, 10);
        $missing = array_diff($required, array_keys($headerMap));
        if ($missing) {
            return $this->json([
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
            ], 400);
        }

        $dataStart = $this->findDataStart($rows, $headerMap);
        if ($dataStart < 0) {
            return $this->json(['error' => 'Unable to find data rows'], 400);
        }

        $db = Database::connect();
        $db->transBegin();

        $inserted = 0;
        $skipped = 0;
        $duplicateCount = 0;
        $duplicates = [];
        $skippedRows = [];
        $skippedLimit = 200;
        $existingKeys = [];
        $seenKeys = [];

        try {
            $existingRows = $db->query('SELECT cid, order_no, start_date, end_date FROM study_leaves')->getResultArray();
            foreach ($existingRows as $row) {
                $key = strtolower(trim((string) ($row['cid'] ?? '')))
                    . '|' . strtolower(trim((string) ($row['order_no'] ?? '')))
                    . '|' . (string) ($row['start_date'] ?? '')
                    . '|' . (string) ($row['end_date'] ?? '');
                $existingKeys[$key] = true;
            }

            $builder = $db->table('study_leaves');
            $dataRows = array_slice($rows, $dataStart);
            foreach ($dataRows as $offset => $row) {
                $rowNumber = $dataStart + $offset + 1;
                $cid = $this->getCell($row, $headerMap, 'cid');
                $fullName = $this->getCell($row, $headerMap, 'full_name');
                if ($cid === null && $fullName === null) {
                    $skipped++;
                    if (count($skippedRows) < $skippedLimit) {
                        $skippedRows[] = [
                            'row' => $rowNumber,
                            'reason' => 'missing cid and full_name',
                            'cid' => $cid ?? '',
                            'full_name' => $fullName ?? '',
                        ];
                    }
                    continue;
                }

                $startDate = $this->parseDateValue($this->getCell($row, $headerMap, 'start_date'));
                $endDate = $this->parseDateValue($this->getCell($row, $headerMap, 'end_date'));
                $approvalDate = null;
                $approvalValue = $this->getCell($row, $headerMap, 'approval_year');
                if ($approvalValue !== null && $this->looksLikeDateCandidate($approvalValue)) {
                    $approvalDate = $this->parseDateValue($approvalValue);
                }
                if ($approvalDate !== null && ($endDate === null || ($startDate !== null && $endDate <= $startDate && $approvalDate >= $startDate))) {
                    $endDate = $approvalDate;
                }
                if ($startDate === null || $endDate === null) {
                    $skipped++;
                    if (count($skippedRows) < $skippedLimit) {
                        $skippedRows[] = [
                            'row' => $rowNumber,
                            'reason' => 'invalid start/end date',
                            'cid' => $cid ?? '',
                            'full_name' => $fullName ?? '',
                            'start_date' => $this->getCell($row, $headerMap, 'start_date'),
                            'end_date' => $this->getCell($row, $headerMap, 'end_date'),
                            'approval_year' => $approvalValue ?? '',
                        ];
                    }
                    continue;
                }

                $programYearsRaw = $this->getCell($row, $headerMap, 'program_years');
                $programYears = (int) preg_replace('/[^0-9]/', '', (string) $programYearsRaw);
                if ($programYears <= 0) {
                    $programYears = 1;
                }

                $orderNo = $this->getCell($row, $headerMap, 'order_no') ?? '';
                $key = strtolower(trim((string) ($cid ?? '')))
                    . '|' . strtolower(trim((string) $orderNo))
                    . '|' . $startDate
                    . '|' . $endDate;
                $isDuplicate = isset($existingKeys[$key]) || isset($seenKeys[$key]);
                if ($isDuplicate) {
                    $duplicateCount++;
                    $skipped++;
                    if (count($duplicates) < 20) {
                        $duplicates[] = [
                            'row' => $rowNumber,
                            'cid' => $cid ?? '',
                            'full_name' => $fullName ?? '',
                            'order_no' => $orderNo,
                            'start_date' => $startDate,
                            'end_date' => $endDate,
                            'source' => isset($existingKeys[$key]) ? 'existing' : 'file',
                        ];
                    }
                    if (count($skippedRows) < $skippedLimit) {
                        $skippedRows[] = [
                            'row' => $rowNumber,
                            'reason' => 'duplicate',
                            'cid' => $cid ?? '',
                            'full_name' => $fullName ?? '',
                            'order_no' => $orderNo,
                            'start_date' => $startDate,
                            'end_date' => $endDate,
                            'source' => isset($existingKeys[$key]) ? 'existing' : 'file',
                        ];
                    }
                    continue;
                }

                $seenKeys[$key] = true;
                $builder->insert([
                    'cid' => $cid ?? '',
                    'full_name' => $fullName ?? '',
                    'position_level' => $this->getCell($row, $headerMap, 'position_level') ?? '',
                    'position_no' => $this->getCell($row, $headerMap, 'position_no') ?? '',
                    'workplace' => $this->getCell($row, $headerMap, 'workplace') ?? '',
                    'program' => $this->getCell($row, $headerMap, 'program') ?? '',
                    'program_years' => $programYears,
                    'institute' => $this->getCell($row, $headerMap, 'institute') ?? '',
                    'start_date' => $startDate,
                    'end_date' => $endDate,
                    'note' => $this->getCell($row, $headerMap, 'note'),
                    'order_no' => $orderNo,
                ]);
                $inserted++;
            }

            $db->transCommit();
        } catch (Throwable $e) {
            $db->transRollback();
            return $this->json(['error' => 'Database error'], 500);
        }

        try {
            $db->table('import_logs')->insert([
                'original_name' => $file->getClientName(),
                'stored_path' => $storedPath,
                'inserted' => $inserted,
                'skipped' => $skipped,
            ]);
        } catch (Throwable $e) {
            // ignore log failure
        }

        return $this->json([
            'success' => true,
            'originalName' => $file->getClientName(),
            'path' => $storedPath,
            'inserted' => $inserted,
            'skipped' => $skipped,
            'duplicate_count' => $duplicateCount,
            'duplicates' => $duplicates,
            'skipped_rows' => $skippedRows,
        ]);
    }

    public function users(): ResponseInterface
    {
        $db = Database::connect();
        $method = strtoupper($this->request->getMethod());

        if ($method === 'GET') {
            try {
                $users = $db->query('SELECT user_id AS id, username, full_name, email, `role`, status, created_at FROM users ORDER BY user_id DESC')->getResultArray();
            } catch (Throwable $e) {
                return $this->json(['error' => 'Database error'], 500);
            }
            return $this->json(['data' => $users]);
        }

        if ($method === 'POST') {
            $payload = $this->readJson();
            $required = ['username', 'password', 'full_name', 'email', 'role'];
            foreach ($required as $field) {
                if (empty($payload[$field])) {
                    return $this->json(['error' => 'Missing field: ' . $field], 400);
                }
            }
            $passwordHash = password_hash((string) $payload['password'], PASSWORD_DEFAULT);
            try {
                $db->table('users')->insert([
                    'username' => $payload['username'],
                    'password' => $passwordHash,
                    'full_name' => $payload['full_name'],
                    'email' => $payload['email'],
                    'role' => $payload['role'],
                    'status' => 'active',
                ]);
            } catch (Throwable $e) {
                return $this->json(['error' => 'Database error'], 500);
            }
            return $this->json(['success' => true]);
        }

        if ($method === 'PUT') {
            $payload = $this->readJson();
            $id = isset($payload['id']) ? (int) $payload['id'] : 0;
            if ($id <= 0) {
                return $this->json(['error' => 'Missing id'], 400);
            }
            $required = ['username', 'full_name', 'email', 'role'];
            foreach ($required as $field) {
                if (empty($payload[$field])) {
                    return $this->json(['error' => 'Missing field: ' . $field], 400);
                }
            }

            $data = [
                'username' => $payload['username'],
                'full_name' => $payload['full_name'],
                'email' => $payload['email'],
                'role' => $payload['role'],
            ];

            if (! empty($payload['password'])) {
                $data['password'] = password_hash((string) $payload['password'], PASSWORD_DEFAULT);
            }

            try {
                $db->table('users')->where('user_id', $id)->update($data);
            } catch (Throwable $e) {
                return $this->json(['error' => 'Database error'], 500);
            }
            return $this->json(['success' => true]);
        }

        if ($method === 'DELETE') {
            $id = (int) $this->request->getGet('id');
            if ($id <= 0) {
                return $this->json(['error' => 'Missing id'], 400);
            }
            try {
                $db->table('users')->where('user_id', $id)->delete();
            } catch (Throwable $e) {
                return $this->json(['error' => 'Database error'], 500);
            }
            return $this->json(['success' => true]);
        }

        return $this->json(['error' => 'Method not allowed'], 405);
    }
}
