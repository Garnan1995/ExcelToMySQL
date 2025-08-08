<?php
/**
 * Test Script to Debug VLOOKUP Dependencies
 * Place this file in your api/ folder
 */

header('Content-Type: text/plain');

// Go up one directory to access config
require_once __DIR__ . '/../config/database.php';
$config = require __DIR__ . '/../config/database.php';

try {
    $pdo = new PDO(
        "mysql:host={$config['host']};dbname={$config['database']};charset={$config['charset']}",
        $config['username'],
        $config['password']
    );
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    
    // Get parameters from URL
    $sheetId = $_GET['sheet_id'] ?? 1;
    $testCell = $_GET['cell'] ?? 'D4';
    
    echo "===========================================\n";
    echo "DEPENDENCY CHAIN ANALYSIS\n";
    echo "===========================================\n\n";
    
    echo "Testing cell: $testCell in sheet $sheetId\n\n";
    
    // 1. Show the current value of the test cell
    echo "1. CURRENT CELL VALUE:\n";
    echo "-------------------------------------------\n";
    
    $stmt = $pdo->prepare("
        SELECT value, value_type FROM cells_independent 
        WHERE sheet_id = ? AND cell_address = ?
        UNION
        SELECT calculated_value as value, 'formula' as value_type FROM cells_dependent 
        WHERE sheet_id = ? AND cell_address = ?
    ");
    $stmt->execute([$sheetId, $testCell, $sheetId, $testCell]);
    $currentValue = $stmt->fetch(PDO::FETCH_ASSOC);
    
    if ($currentValue) {
        echo "Value: {$currentValue['value']}\n";
        echo "Type: {$currentValue['value_type']}\n";
    } else {
        echo "Cell is empty or not found\n";
    }
    
    // 2. Find all formulas that reference this cell
    echo "\n2. FORMULAS THAT DIRECTLY REFERENCE $testCell:\n";
    echo "-------------------------------------------\n";
    
    $cleanCell = str_replace('$', '', $testCell);
    
    $stmt = $pdo->prepare("
        SELECT cell_address, formula, formula_type
        FROM cells_dependent 
        WHERE sheet_id = ? 
        AND (
            formula LIKE ? 
            OR formula LIKE ?
            OR formula LIKE ?
        )
    ");
    
    $stmt->execute([
        $sheetId,
        '%' . $testCell . '%',
        '%' . $cleanCell . '%',
        '%$' . $cleanCell . '%'
    ]);
    
    $dependentFormulas = $stmt->fetchAll(PDO::FETCH_ASSOC);
    
    if ($dependentFormulas) {
        foreach ($dependentFormulas as $formula) {
            echo "\nCell: {$formula['cell_address']}\n";
            echo "Formula: {$formula['formula']}\n";
            echo "Type: {$formula['formula_type']}\n";
        }
    } else {
        echo "No formulas directly reference this cell.\n";
    }
    
    // 3. Check for VLOOKUP formulas that use this cell as lookup value
    echo "\n3. VLOOKUP FORMULAS USING $testCell:\n";
    echo "-------------------------------------------\n";
    
    $stmt = $pdo->prepare("
        SELECT cell_address, formula 
        FROM cells_dependent 
        WHERE sheet_id = ? 
        AND formula_type = 'VLOOKUP'
        AND (
            formula LIKE ?
            OR formula LIKE ?
            OR formula LIKE ?
        )
    ");
    
    $stmt->execute([
        $sheetId,
        '%(' . $testCell . ',%',
        '%(' . $cleanCell . ',%',
        '%($' . $cleanCell . ',%'
    ]);
    
    $vlookups = $stmt->fetchAll(PDO::FETCH_ASSOC);
    
    if ($vlookups) {
        foreach ($vlookups as $vlookup) {
            echo "\nCell: {$vlookup['cell_address']}\n";
            echo "Formula: {$vlookup['formula']}\n";
            echo "*** THIS VLOOKUP USES $testCell AS LOOKUP VALUE ***\n";
        }
    } else {
        echo "No VLOOKUP formulas use this cell as lookup value.\n";
    }
    
    // 4. Show all VLOOKUP formulas in sheet for reference
    echo "\n4. ALL VLOOKUP FORMULAS IN SHEET:\n";
    echo "-------------------------------------------\n";
    
    $stmt = $pdo->prepare("
        SELECT cell_address, formula 
        FROM cells_dependent 
        WHERE sheet_id = ? 
        AND formula_type = 'VLOOKUP'
        LIMIT 10
    ");
    $stmt->execute([$sheetId]);
    
    $allVlookups = $stmt->fetchAll(PDO::FETCH_ASSOC);
    
    foreach ($allVlookups as $vlookup) {
        echo "\nCell: {$vlookup['cell_address']}\n";
        echo "Formula: " . substr($vlookup['formula'], 0, 100) . "...\n";
    }
    
    // 5. Check data validation
    echo "\n5. DATA VALIDATION ON $testCell:\n";
    echo "-------------------------------------------\n";
    
    $stmt = $pdo->prepare("
        SELECT * FROM data_validations 
        WHERE sheet_id = ? 
        AND (
            cell_range = ?
            OR cell_range LIKE ?
            OR cell_range LIKE ?
        )
    ");
    
    $stmt->execute([
        $sheetId,
        $testCell,
        $testCell . ':%',
        '%:' . $testCell
    ]);
    
    $validations = $stmt->fetchAll(PDO::FETCH_ASSOC);
    
    if ($validations) {
        foreach ($validations as $validation) {
            echo "Range: {$validation['cell_range']}\n";
            echo "Type: {$validation['validation_type']}\n";
            if ($validation['validation_formula']) {
                echo "Formula: {$validation['validation_formula']}\n";
            }
            if ($validation['validation_list']) {
                echo "List: {$validation['validation_list']}\n";
            }
        }
    } else {
        echo "No data validation found on this cell.\n";
    }
    
    echo "\n===========================================\n";
    echo "END OF ANALYSIS\n";
    echo "===========================================\n";
    
} catch (Exception $e) {
    echo "ERROR: " . $e->getMessage() . "\n";
    echo "Stack trace:\n" . $e->getTraceAsString() . "\n";
}
?>