<?php
/**
 * phpMyAdmin UI Helper - Creates a user-friendly interface
 * Save this as: src/ExcelDatabaseUI.php
 */

class ExcelDatabaseUI {
    private $pdo;
    
    public function __construct($pdo) {
        $this->pdo = $pdo;
    }
    
    /**
     * Generate HTML table view of sheet data
     */
    public function generateSheetView($sheetId) {
        $html = '<style>
            .excel-grid { border-collapse: collapse; font-family: Arial; }
            .excel-grid th, .excel-grid td { 
                border: 1px solid #ddd; 
                padding: 4px 8px; 
                min-width: 60px;
            }
            .excel-grid th { background: #f0f0f0; position: sticky; top: 0; }
            .cell-formula { background: #e8f4ff; }
            .cell-value { background: #fff; }
            .cell-error { background: #ffe8e8; color: red; }
            .cell-dropdown { background: #fffacd; }
            .cell-address { background: #f0f0f0; font-weight: bold; }
        </style>';
        
        // Get sheet info
        $stmt = $this->pdo->prepare("
            SELECT sheet_name, row_count, column_count 
            FROM sheets WHERE id = ?
        ");
        $stmt->execute([$sheetId]);
        $sheetInfo = $stmt->fetch(PDO::FETCH_ASSOC);
        
        $html .= "<h2>Sheet: {$sheetInfo['sheet_name']}</h2>";
        $html .= '<table class="excel-grid">';
        
        // Generate header row with column letters
        $html .= '<tr><th></th>';
        for ($col = 1; $col <= min($sheetInfo['column_count'], 26); $col++) {
            $html .= '<th>' . chr(64 + $col) . '</th>';
        }
        $html .= '</tr>';
        
        // Get all cells for this sheet
        $sql = "
            SELECT 
                'independent' as type,
                cell_address, row_num, col_num, value, 
                NULL as formula, comment
            FROM cells_independent WHERE sheet_id = ?
            UNION ALL
            SELECT 
                'dependent' as type,
                cell_address, row_num, col_num, 
                calculated_value as value, formula, comment
            FROM cells_dependent WHERE sheet_id = ?
            ORDER BY row_num, col_num
        ";
        
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute([$sheetId, $sheetId]);
        
        $cells = [];
        while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
            $cells[$row['row_num']][$row['col_num']] = $row;
        }
        
        // Generate rows
        for ($row = 1; $row <= min($sheetInfo['row_count'], 100); $row++) {
            $html .= "<tr><td class='cell-address'>$row</td>";
            
            for ($col = 1; $col <= min($sheetInfo['column_count'], 26); $col++) {
                if (isset($cells[$row][$col])) {
                    $cell = $cells[$row][$col];
                    $class = $cell['type'] == 'dependent' ? 'cell-formula' : 'cell-value';
                    $title = $cell['formula'] ? "Formula: {$cell['formula']}" : '';
                    $comment = $cell['comment'] ? " 💬" : '';
                    
                    $html .= "<td class='$class' title='$title'>";
                    $html .= htmlspecialchars($cell['value']) . $comment;
                    $html .= "</td>";
                } else {
                    $html .= "<td></td>";
                }
            }
            $html .= '</tr>';
        }
        
        $html .= '</table>';
        return $html;
    }
    
    /**
     * Generate formula dependency graph
     */
    public function generateDependencyGraph($sheetId) {
        $sql = "
            SELECT 
                cd.cell_address as source,
                cd.dependencies
            FROM cells_dependent cd
            WHERE cd.sheet_id = ? 
            AND cd.dependencies IS NOT NULL 
            AND cd.dependencies != '[]'
        ";
        
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute([$sheetId]);
        
        $graph = [];
        while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
            $deps = json_decode($row['dependencies'], true);
            $graph[$row['source']] = $deps;
        }
        
        return $graph;
    }
    
    /**
     * Generate edit form for cell
     */
    public function generateCellEditForm($sheetId, $cellAddress) {
        $stmt = $this->pdo->prepare("
            SELECT * FROM cells_independent 
            WHERE sheet_id = ? AND cell_address = ?
            UNION
            SELECT * FROM cells_dependent 
            WHERE sheet_id = ? AND cell_address = ?
        ");
        $stmt->execute([$sheetId, $cellAddress, $sheetId, $cellAddress]);
        $cell = $stmt->fetch(PDO::FETCH_ASSOC);
        
        $html = '<form method="POST" action="update_cell.php">';
        $html .= '<h3>Edit Cell: ' . $cellAddress . '</h3>';
        $html .= '<input type="hidden" name="sheet_id" value="' . $sheetId . '">';
        $html .= '<input type="hidden" name="cell_address" value="' . $cellAddress . '">';
        
        if (isset($cell['formula'])) {
            $html .= '<label>Formula:</label><br>';
            $html .= '<textarea name="formula" rows="3" cols="50">' . 
                     htmlspecialchars($cell['formula']) . '</textarea><br>';
            $html .= '<label>Calculated Value:</label><br>';
            $html .= '<input type="text" value="' . 
                     htmlspecialchars($cell['calculated_value']) . '" readonly><br>';
        } else {
            $html .= '<label>Value:</label><br>';
            $html .= '<input type="text" name="value" value="' . 
                     htmlspecialchars($cell['value'] ?? '') . '"><br>';
        }
        
        $html .= '<label>Comment:</label><br>';
        $html .= '<textarea name="comment" rows="2" cols="50">' . 
                 htmlspecialchars($cell['comment'] ?? '') . '</textarea><br>';
        
        $html .= '<button type="submit">Update Cell</button>';
        $html .= '</form>';
        
        return $html;
    }
    
    /**
     * Generate a summary dashboard for a workbook
     */
    public function generateWorkbookDashboard($workbookId) {
        $html = '<style>
            .dashboard { font-family: Arial, sans-serif; }
            .stats-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin: 20px 0; }
            .stat-card { 
                background: #f9f9f9; 
                border: 1px solid #ddd; 
                border-radius: 8px; 
                padding: 15px;
            }
            .stat-card h3 { margin-top: 0; color: #333; }
            .stat-value { font-size: 24px; font-weight: bold; color: #0066cc; }
            .sheet-list { list-style: none; padding: 0; }
            .sheet-list li { 
                background: #fff; 
                border: 1px solid #eee; 
                padding: 10px; 
                margin: 5px 0;
                border-radius: 4px;
            }
            .formula-stats { background: #e8f4ff; padding: 10px; border-radius: 4px; }
        </style>';
        
        $html .= '<div class="dashboard">';
        
        // Get workbook info
        $stmt = $this->pdo->prepare("
            SELECT * FROM workbooks WHERE id = ?
        ");
        $stmt->execute([$workbookId]);
        $workbook = $stmt->fetch(PDO::FETCH_ASSOC);
        
        $html .= "<h1>Workbook: {$workbook['filename']}</h1>";
        $html .= "<p>Imported: {$workbook['import_date']}</p>";
        
        // Statistics
        $html .= '<div class="stats-grid">';
        
        // Sheet count
        $stmt = $this->pdo->prepare("
            SELECT COUNT(*) as count FROM sheets WHERE workbook_id = ?
        ");
        $stmt->execute([$workbookId]);
        $sheetCount = $stmt->fetchColumn();
        
        $html .= '<div class="stat-card">';
        $html .= '<h3>Total Sheets</h3>';
        $html .= '<div class="stat-value">' . $sheetCount . '</div>';
        $html .= '</div>';
        
        // Total cells
        $stmt = $this->pdo->prepare("
            SELECT 
                (SELECT COUNT(*) FROM cells_independent ci 
                 JOIN sheets s ON ci.sheet_id = s.id 
                 WHERE s.workbook_id = ?) +
                (SELECT COUNT(*) FROM cells_dependent cd 
                 JOIN sheets s ON cd.sheet_id = s.id 
                 WHERE s.workbook_id = ?) as total
        ");
        $stmt->execute([$workbookId, $workbookId]);
        $totalCells = $stmt->fetchColumn();
        
        $html .= '<div class="stat-card">';
        $html .= '<h3>Total Cells</h3>';
        $html .= '<div class="stat-value">' . number_format($totalCells) . '</div>';
        $html .= '</div>';
        
        // Formula count
        $stmt = $this->pdo->prepare("
            SELECT COUNT(*) FROM cells_dependent cd
            JOIN sheets s ON cd.sheet_id = s.id
            WHERE s.workbook_id = ?
        ");
        $stmt->execute([$workbookId]);
        $formulaCount = $stmt->fetchColumn();
        
        $html .= '<div class="stat-card">';
        $html .= '<h3>Formulas</h3>';
        $html .= '<div class="stat-value">' . number_format($formulaCount) . '</div>';
        $html .= '</div>';
        
        $html .= '</div>'; // End stats-grid
        
        // Sheet list with details
        $html .= '<h2>Sheets</h2>';
        $html .= '<ul class="sheet-list">';
        
        $stmt = $this->pdo->prepare("
            SELECT s.*,
                   (SELECT COUNT(*) FROM cells_independent WHERE sheet_id = s.id) as value_cells,
                   (SELECT COUNT(*) FROM cells_dependent WHERE sheet_id = s.id) as formula_cells
            FROM sheets s
            WHERE s.workbook_id = ?
            ORDER BY s.sheet_index
        ");
        $stmt->execute([$workbookId]);
        
        while ($sheet = $stmt->fetch(PDO::FETCH_ASSOC)) {
            $html .= '<li>';
            $html .= '<strong>' . htmlspecialchars($sheet['sheet_name']) . '</strong><br>';
            $html .= "Size: {$sheet['row_count']} × {$sheet['column_count']}<br>";
            $html .= "Cells: {$sheet['value_cells']} values, {$sheet['formula_cells']} formulas<br>";
            $html .= '<a href="view_workbook.php?sheet_id=' . $sheet['id'] . '">View Sheet</a>';
            $html .= '</li>';
        }
        
        $html .= '</ul>';
        
        // Formula statistics
        $html .= '<h2>Formula Analysis</h2>';
        $html .= '<div class="formula-stats">';
        
        $stmt = $this->pdo->prepare("
            SELECT formula_type, COUNT(*) as count
            FROM cells_dependent cd
            JOIN sheets s ON cd.sheet_id = s.id
            WHERE s.workbook_id = ?
            GROUP BY formula_type
            ORDER BY count DESC
        ");
        $stmt->execute([$workbookId]);
        
        $html .= '<h3>Formula Types Used:</h3>';
        $html .= '<ul>';
        while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
            $html .= '<li>' . htmlspecialchars($row['formula_type']) . ': ' . $row['count'] . '</li>';
        }
        $html .= '</ul>';
        
        $html .= '</div>';
        $html .= '</div>'; // End dashboard
        
        return $html;
    }
    
    /**
     * Generate a view of all VLOOKUP formulas
     */
    public function generateVLOOKUPAnalysis($workbookId) {
        $html = '<style>
            .vlookup-table { 
                width: 100%; 
                border-collapse: collapse; 
                font-family: monospace;
                font-size: 12px;
            }
            .vlookup-table th { 
                background: #0066cc; 
                color: white; 
                padding: 8px;
                text-align: left;
            }
            .vlookup-table td { 
                border: 1px solid #ddd; 
                padding: 6px;
            }
            .formula-cell { background: #f0f8ff; }
            .error-cell { background: #ffeeee; color: red; }
            .success-cell { background: #eeffee; }
        </style>';
        
        $html .= '<h2>VLOOKUP Formula Analysis</h2>';
        
        $stmt = $this->pdo->prepare("
            SELECT 
                s.sheet_name,
                cd.cell_address,
                cd.formula,
                cd.calculated_value,
                cd.has_error,
                cd.error_message,
                cd.dependencies,
                cd.external_references
            FROM cells_dependent cd
            JOIN sheets s ON cd.sheet_id = s.id
            WHERE s.workbook_id = ? AND cd.formula_type = 'VLOOKUP'
            ORDER BY s.sheet_index, cd.row_num, cd.col_num
        ");
        $stmt->execute([$workbookId]);
        
        $html .= '<table class="vlookup-table">';
        $html .= '<tr>';
        $html .= '<th>Sheet</th>';
        $html .= '<th>Cell</th>';
        $html .= '<th>Formula</th>';
        $html .= '<th>Result</th>';
        $html .= '<th>Dependencies</th>';
        $html .= '<th>Status</th>';
        $html .= '</tr>';
        
        while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
            $statusClass = $row['has_error'] ? 'error-cell' : 'success-cell';
            $status = $row['has_error'] ? $row['error_message'] : 'OK';
            
            $html .= '<tr>';
            $html .= '<td>' . htmlspecialchars($row['sheet_name']) . '</td>';
            $html .= '<td class="formula-cell">' . htmlspecialchars($row['cell_address']) . '</td>';
            $html .= '<td>' . htmlspecialchars($row['formula']) . '</td>';
            $html .= '<td>' . htmlspecialchars($row['calculated_value']) . '</td>';
            $html .= '<td>' . htmlspecialchars($row['dependencies']) . '</td>';
            $html .= '<td class="' . $statusClass . '">' . $status . '</td>';
            $html .= '</tr>';
        }
        
        $html .= '</table>';
        
        return $html;
    }
}