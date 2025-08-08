<?php
header('Content-Type: application/json');
error_reporting(E_ALL);
ini_set('display_errors', 0); // Don't display errors as HTML

try {
    // Get JSON input
    $input = json_decode(file_get_contents('php://input'), true);
    
    if (!$input) {
        throw new Exception('Invalid input data');
    }
    
    // Include required files
    require_once __DIR__ . '/../config/database.php';
    require_once __DIR__ . '/../src/ExcelFormulaProcessor.php';
    
    $config = require __DIR__ . '/../config/database.php';
    
    $pdo = new PDO(
        "mysql:host={$config['host']};dbname={$config['database']};charset={$config['charset']}",
        $config['username'],
        $config['password']
    );
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    
    // For now, let's do a simple update without the formula processor
    // to test if the basic update works
    
    $sheetId = $input['sheet_id'];
    $cellAddress = $input['cell_address'];
    $value = $input['value'];
    $isFormula = $input['is_formula'] ?? false;
    
    // Parse cell address
    preg_match('/([A-Z]+)(\d+)/', $cellAddress, $matches);
    $colLetter = $matches[1] ?? '';
    $rowNum = $matches[2] ?? 0;
    $colNum = 0;
    
    // Convert column letter to number
    for ($i = 0; $i < strlen($colLetter); $i++) {
        $colNum = $colNum * 26 + (ord($colLetter[$i]) - 64);
    }
    
    if ($isFormula) {
        // Handle formula cell
        $stmt = $pdo->prepare("
            INSERT INTO cells_dependent 
            (sheet_id, cell_address, row_num, col_num, col_letter, formula, calculated_value, formula_type)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON DUPLICATE KEY UPDATE
            formula = VALUES(formula),
            calculated_value = VALUES(calculated_value)
        ");
        
        // For now, just store the formula without evaluation
        $calculatedValue = $value; // This should be calculated
        $formulaType = 'CUSTOM';
        
        $stmt->execute([
            $sheetId, $cellAddress, $rowNum, $colNum, $colLetter,
            $value, $calculatedValue, $formulaType
        ]);
    } else {
        // First check if this cell exists in dependent cells and remove it
        $stmt = $pdo->prepare("DELETE FROM cells_dependent WHERE sheet_id = ? AND cell_address = ?");
        $stmt->execute([$sheetId, $cellAddress]);
        
        // Handle regular value cell
        $stmt = $pdo->prepare("
            INSERT INTO cells_independent 
            (sheet_id, cell_address, row_num, col_num, col_letter, value, value_type, formatted_value)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON DUPLICATE KEY UPDATE
            value = VALUES(value),
            formatted_value = VALUES(formatted_value),
            updated_at = CURRENT_TIMESTAMP
        ");
        
        $valueType = is_numeric($value) ? 'number' : 'string';
        
        $stmt->execute([
            $sheetId, $cellAddress, $rowNum, $colNum, $colLetter,
            $value, $valueType, $value
        ]);
    }
    
    // Get dependent cells that need recalculation
    $stmt = $pdo->prepare("
        SELECT cell_address 
        FROM cells_dependent 
        WHERE sheet_id = ? 
        AND dependencies LIKE ?
    ");
    $stmt->execute([$sheetId, '%"' . $cellAddress . '"%']);
    
    $affectedCells = $stmt->fetchAll(PDO::FETCH_COLUMN);
    
    // Return success response
    echo json_encode([
        'success' => true,
        'cell' => [
            'cell_address' => $cellAddress,
            'value' => $value,
            'formula' => $isFormula ? $value : null,
            'formatted_value' => $value
        ],
        'affected_cells' => $affectedCells
    ]);
    
} catch (Exception $e) {
    // Return error as JSON, not HTML
    http_response_code(500);
    echo json_encode([
        'success' => false,
        'error' => $e->getMessage(),
        'trace' => $e->getTraceAsString()
    ]);
}