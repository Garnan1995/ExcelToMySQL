<?php
header('Content-Type: application/json');
error_reporting(E_ALL);
ini_set('display_errors', 0);

try {
    // Get JSON input
    $input = json_decode(file_get_contents('php://input'), true);
    
    if (!$input) {
        throw new Exception('Invalid input data');
    }
    
    require_once __DIR__ . '/../config/database.php';
    $config = require __DIR__ . '/../config/database.php';
    
    $pdo = new PDO(
        "mysql:host={$config['host']};dbname={$config['database']};charset={$config['charset']}",
        $config['username'],
        $config['password']
    );
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    
    $sheetId = $input['sheet_id'];
    $cellsToRecalculate = $input['cells'] ?? [];
    
    $results = [];
    
    // For each cell that needs recalculation
    foreach ($cellsToRecalculate as $cellAddress) {
        // Get the formula for this cell
        $stmt = $pdo->prepare("
            SELECT formula, dependencies 
            FROM cells_dependent 
            WHERE sheet_id = ? AND cell_address = ?
        ");
        $stmt->execute([$sheetId, $cellAddress]);
        $cellData = $stmt->fetch(PDO::FETCH_ASSOC);
        
        if ($cellData && $cellData['formula']) {
            // Simple recalculation - evaluate VLOOKUP
            $calculatedValue = recalculateFormula($pdo, $sheetId, $cellData['formula'], $cellData['dependencies']);
            
            // Update the calculated value
            $updateStmt = $pdo->prepare("
                UPDATE cells_dependent 
                SET calculated_value = ?, 
                    formatted_value = ?,
                    has_error = ?,
                    error_message = ?
                WHERE sheet_id = ? AND cell_address = ?
            ");
            
            $hasError = (strpos($calculatedValue, '#') === 0);
            $errorMsg = $hasError ? $calculatedValue : null;
            
            $updateStmt->execute([
                $calculatedValue,
                $calculatedValue,
                $hasError ? 1 : 0,
                $errorMsg,
                $sheetId,
                $cellAddress
            ]);
            
            $results[] = [
                'cell_address' => $cellAddress,
                'value' => $calculatedValue,
                'formula' => $cellData['formula'],
                'formatted_value' => $calculatedValue,
                'error_message' => $errorMsg
            ];
        }
    }
    
    echo json_encode($results);
    
} catch (Exception $e) {
    http_response_code(500);
    echo json_encode([
        'error' => $e->getMessage()
    ]);
}

/**
 * Simple formula recalculation
 */
function recalculateFormula($pdo, $sheetId, $formula, $dependencies) {
    // Remove the = sign
    $formula = ltrim($formula, '=');
    
    // Handle VLOOKUP - simplified version
    if (stripos($formula, 'VLOOKUP') !== false) {
        // Parse VLOOKUP(lookup_value, table_array, col_index, exact_match)
        if (preg_match('/VLOOKUP\s*\(\s*([^,]+),\s*([^,]+),\s*([^,]+),\s*([^)]+)\)/i', $formula, $matches)) {
            $lookupRef = trim($matches[1], ' $');
            $tableRange = trim($matches[2]);
            $colIndex = trim($matches[3]);
            $exactMatch = trim($matches[4]);
            
            // Get lookup value
            $lookupValue = getCellValue($pdo, $sheetId, $lookupRef);
            
            // Parse table range (e.g., 'Standar Durasi'!$B$5:$AM$83)
            if (preg_match("/(?:'([^']+)'!)?\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/", $tableRange, $rangeMatch)) {
                $sheetName = $rangeMatch[1];
                $startCol = $rangeMatch[2];
                $startRow = $rangeMatch[3];
                $endCol = $rangeMatch[4];
                $endRow = $rangeMatch[5];
                
                // Get the sheet ID for the lookup table
                $targetSheetId = $sheetId;
                if ($sheetName) {
                    $stmt = $pdo->prepare("
                        SELECT id FROM sheets 
                        WHERE sheet_name = ? 
                        AND workbook_id = (
                            SELECT workbook_id FROM sheets WHERE id = ?
                        )
                    ");
                    $stmt->execute([$sheetName, $sheetId]);
                    $targetSheetId = $stmt->fetchColumn();
                }
                
                // Perform the lookup - find the row with matching value in first column
                $stmt = $pdo->prepare("
                    SELECT row_num 
                    FROM cells_independent 
                    WHERE sheet_id = ? 
                    AND col_letter = ? 
                    AND value = ?
                    AND row_num BETWEEN ? AND ?
                    LIMIT 1
                ");
                $stmt->execute([$targetSheetId, $startCol, $lookupValue, $startRow, $endRow]);
                $matchRow = $stmt->fetchColumn();
                
                if ($matchRow) {
                    // Calculate target column for return value
                    // If colIndex is a MATCH function, evaluate it
                    if (stripos($colIndex, 'MATCH') !== false) {
                        // For now, just use a default column index
                        $colIndex = 2;
                    } else {
                        $colIndex = intval($colIndex);
                    }
                    
                    // Get the target column letter
                    $targetCol = chr(ord($startCol) + $colIndex - 1);
                    
                    // Get the value from the target cell
                    $stmt = $pdo->prepare("
                        SELECT value 
                        FROM cells_independent 
                        WHERE sheet_id = ? 
                        AND col_letter = ? 
                        AND row_num = ?
                    ");
                    $stmt->execute([$targetSheetId, $targetCol, $matchRow]);
                    $result = $stmt->fetchColumn();
                    
                    return $result ?: '0';
                }
            }
            
            return '#N/A';
        }
    }
    
    // Handle IFERROR
    if (stripos($formula, 'IFERROR') !== false) {
        // For now, just return 0 for IFERROR
        return '0';
    }
    
    // Handle simple cell references
    if (preg_match('/^([A-Z]+\d+)$/', $formula, $matches)) {
        return getCellValue($pdo, $sheetId, $matches[1]);
    }
    
    // Default return
    return $formula;
}

/**
 * Get cell value helper
 */
function getCellValue($pdo, $sheetId, $cellAddress) {
    // Remove $ signs
    $cellAddress = str_replace('$', '', $cellAddress);
    
    $stmt = $pdo->prepare("
        SELECT value FROM cells_independent 
        WHERE sheet_id = ? AND cell_address = ?
        UNION
        SELECT calculated_value FROM cells_dependent 
        WHERE sheet_id = ? AND cell_address = ?
        LIMIT 1
    ");
    $stmt->execute([$sheetId, $cellAddress, $sheetId, $cellAddress]);
    
    $value = $stmt->fetchColumn();
    return $value !== false ? $value : '';
}