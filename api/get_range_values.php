<?php
/**
 * Get Range Values API
 * Returns all values from a specified range (for dropdown lists)
 */

header('Content-Type: application/json');
header('Access-Control-Allow-Origin: *');

require_once '../config/database.php';
$config = require '../config/database.php';

try {
    $pdo = new PDO(
        "mysql:host={$config['host']};dbname={$config['database']};charset={$config['charset']}",
        $config['username'],
        $config['password']
    );
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    
    $sheetName = $_GET['sheet'] ?? '';
    $range = $_GET['range'] ?? '';
    
    error_log("Getting range values for sheet: $sheetName, range: $range");
    
    // Remove $ signs and quotes from range
    $range = str_replace('$', '', $range);
    $range = str_replace("'", '', $range);
    
    // Parse range (e.g., R3:R42 or B5:B83)
    if (!preg_match('/([A-Z]+)(\d+):([A-Z]+)(\d+)/', $range, $matches)) {
        throw new Exception("Invalid range format: $range");
    }
    
    $startCol = $matches[1];
    $startRow = intval($matches[2]);
    $endCol = $matches[3];
    $endRow = intval($matches[4]);
    
    error_log("Parsed range: $startCol$startRow:$endCol$endRow");
    
    // Get sheet ID from sheet name
    $sheetId = null;
    
    if ($sheetName) {
        $stmt = $pdo->prepare("SELECT id FROM sheets WHERE sheet_name = ?");
        $stmt->execute([$sheetName]);
        $sheetId = $stmt->fetchColumn();
        
        error_log("Found sheet ID: $sheetId for sheet: $sheetName");
    }
    
    if (!$sheetId) {
        // If no sheet name provided or not found, try to get from the current context
        // This is a fallback - you might want to require sheet name
        $stmt = $pdo->query("SELECT id FROM sheets ORDER BY id DESC LIMIT 1");
        $sheetId = $stmt->fetchColumn();
        
        error_log("Using fallback sheet ID: $sheetId");
    }
    
    if ($sheetId) {
        // Get all values from the range
        // For dropdown lists, we typically want values from a single column
        $values = [];
        
        // If it's a single column range
        if ($startCol === $endCol) {
            $stmt = $pdo->prepare("
                SELECT DISTINCT value 
                FROM cells_independent 
                WHERE sheet_id = ? 
                AND col_letter = ? 
                AND row_num BETWEEN ? AND ?
                AND value IS NOT NULL 
                AND value != ''
                
                UNION
                
                SELECT DISTINCT calculated_value as value
                FROM cells_dependent
                WHERE sheet_id = ?
                AND col_letter = ?
                AND row_num BETWEEN ? AND ?
                AND calculated_value IS NOT NULL
                AND calculated_value != ''
                
                ORDER BY value
            ");
            
            $stmt->execute([
                $sheetId, $startCol, $startRow, $endRow,
                $sheetId, $startCol, $startRow, $endRow
            ]);
            
            $values = $stmt->fetchAll(PDO::FETCH_COLUMN);
            
            error_log("Found " . count($values) . " values in range");
            
            // Remove duplicates and empty values
            $values = array_filter($values, function($v) {
                return $v !== null && $v !== '';
            });
            
            $values = array_unique($values);
            $values = array_values($values); // Re-index array
            
        } else {
            // Multi-column range - return as 2D array
            for ($row = $startRow; $row <= $endRow; $row++) {
                $rowValues = [];
                
                for ($col = columnToNumber($startCol); $col <= columnToNumber($endCol); $col++) {
                    $colLetter = numberToColumn($col);
                    
                    $stmt = $pdo->prepare("
                        SELECT value FROM cells_independent 
                        WHERE sheet_id = ? AND col_letter = ? AND row_num = ?
                        
                        UNION
                        
                        SELECT calculated_value as value FROM cells_dependent
                        WHERE sheet_id = ? AND col_letter = ? AND row_num = ?
                        
                        LIMIT 1
                    ");
                    
                    $stmt->execute([
                        $sheetId, $colLetter, $row,
                        $sheetId, $colLetter, $row
                    ]);
                    
                    $value = $stmt->fetchColumn();
                    if ($value !== false) {
                        $rowValues[] = $value;
                    }
                }
                
                if (!empty($rowValues)) {
                    // For dropdown purposes, we might want just the first column
                    $values[] = $rowValues[0];
                }
            }
        }
        
        echo json_encode([
            'success' => true,
            'values' => $values,
            'sheet_id' => $sheetId,
            'range' => "$startCol$startRow:$endCol$endRow"
        ]);
        
    } else {
        echo json_encode([
            'success' => false,
            'values' => [],
            'error' => 'Sheet not found'
        ]);
    }
    
} catch (Exception $e) {
    error_log("Error in get_range_values.php: " . $e->getMessage());
    
    echo json_encode([
        'success' => false,
        'error' => $e->getMessage(),
        'values' => []
    ]);
}

/**
 * Convert column letter to number
 */
function columnToNumber($col) {
    $num = 0;
    for ($i = 0; $i < strlen($col); $i++) {
        $num = $num * 26 + (ord($col[$i]) - 64);
    }
    return $num;
}

/**
 * Convert number to column letter
 */
function numberToColumn($num) {
    $col = '';
    while ($num > 0) {
        $num--;
        $col = chr(65 + ($num % 26)) . $col;
        $num = intval($num / 26);
    }
    return $col;
}