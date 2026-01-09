<?php

namespace App\Database\Migrations;

use CodeIgniter\Database\Migration;

class AddPositionPartsToStudyLeaves extends Migration
{
    public function up()
    {
        $fields = [
            'position_title' => [
                'type' => 'VARCHAR',
                'constraint' => 255,
                'null' => true,
                'after' => 'position_level',
            ],
            'position_hospital' => [
                'type' => 'VARCHAR',
                'constraint' => 255,
                'null' => true,
                'after' => 'position_title',
            ],
            'position_office' => [
                'type' => 'VARCHAR',
                'constraint' => 255,
                'null' => true,
                'after' => 'position_hospital',
            ],
        ];

        $this->forge->addColumn('study_leaves', $fields);
    }

    public function down()
    {
        $this->forge->dropColumn('study_leaves', ['position_title', 'position_hospital', 'position_office']);
    }
}
