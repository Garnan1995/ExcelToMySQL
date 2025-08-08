<?php
/**
 * Fixed Cell Update API with Proper VLOOKUP Dependency Resolution
 * This version correctly identifies and updates all dependent formulas
 */

header('Content-Type: application/json');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    exit(0);
}

error_reporting(E_ALL);
ini_set('display_errors', 0);
ini_set('max_execution_time', 60);

try {
    // Get JSON input
    $input = json_decode(file_get_contents('php://input'), true);
    
    if (!$input) {
        throw new Exception('Invalid input data');
    }
    
    // Database connection
    require_once __DIR__ . '/../config/database.php';
    $config = require __DIR__ . '/../config/database.php';
    
    $pdo = new PDO(
        "mysql:host={$config['host']};dbname={$config['database']};charset={$config['charset']}",
        $config['username'],
        $config['password']
    );
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    
    // Start transaction
    $pdo->beginTransaction();
    
    // Extract parameters
    $sheetId = $input['sheet_id'];
    $cellAddress = trim($input['cell_address']);
    $value = $input['value'];
    $isFormula = isset($input['is_formula']) ? $input['is_formula'] : (substr($value, 0, 1) === '=');
    
    // Parse cell address
    if (!preg_match('/([A-Z]+)(\d+)/', $cellAddress, $matches)) {
        throw new Exception('Invalid cell address format');
    }
    
    $colLetter = $matches[1];
    $rowNum = intval($matches[2]);
    $colNum = columnLetterToNumber($colLetter);
    
    // Debug logging
    error_log("Updating cell: $cellAddress with value: $value");
    
    // Store the update
    if ($isFormula) {
        updateFormulaCell($pdo, $sheetId, $cellAddress, $rowNum, $colNum, $colLetter, $value);
    } else {
        updateValueCell($pdo, $sheetId, $cellAddress, $rowNum, $colNum, $colLetter, $value);
    }
    
    // Commit the direct update first
    $pdo->commit();
    
    // Now find ALL cells that depend on this cell (including VLOOKUP formulas)
    $affectedCells = findAllDependentCells($pdo, $sheetId, $cellAddress);
    
    error_log("Found " . count($affectedCells) . " dependent cells for $cellAddress");
    error_log("Dependent cells: " . json_encode($affectedCells));
    
    // Return the response
    echo json_encode([
        'success' => true,
        'cell' => [
            'cell_address' => $cellAddress,
            'value' => $value,
            'formula' => $isFormula ? $value : null,
            'formatted_value' => $value,
            'value_type' => $isFormula ? 'formula' : determineValueType($value)
        ],
        'affected_cells' => $affectedCells,
        'message' => 'Cell updated successfully'
    ]);
    
} catch (Exception $e) {
    if (isset($pdo) && $pdo->inTransaction()) {
        $pdo->rollBack();
    }
    
    error_log("Error in update_cell.php: " . $e->getMessage());
    
    http_response_code(500);
    echo json_encode([
        'success' => false,
        'error' => $e->getMessage()
    ]);
}

/**
 * Update a formula cell
 */
function updateFormulaCell($pdo, $sheetId, $cellAddress, $rowNum, $colNum, $colLetter, $formula) {
    // Extract formula type and dependencies
    $formulaType = extractFormulaType($formula);
    $dependencies = extractDependencies($formula);
    
    // For now, store the formula as-is (calculation happens in recalculate.php)
    $stmt = $pdo->prepare("
        INSERT INTO cells_dependent 
        (sheet_id, cell_address, row_num, col_num, col_letter, 
         formula, formula_type, calculated_value, formatted_value, dependencies)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON DUPLICATE KEY UPDATE
            formula = VALUES(formula),
            formula_type = VALUES(formula_type),
            calculated_value = VALUES(calculated_value),
            formatted_value = VALUES(formatted_value),
            dependencies = VALUES(dependencies),
            updated_at = CURRENT_TIMESTAMP
    ");
    
    $stmt->execute([
        $sheetId, $cellAddress, $rowNum, $colNum, $colLetter,
        $formula, $formulaType, '', '', json_encode($dependencies)
    ]);
    
    // Remove from independent cells if exists
    $stmt = $pdo->prepare("DELETE FROM cells_independent WHERE sheet_id = ? AND cell_address = ?");
    $stmt->execute([$sheetId, $cellAddress]);
}

/**
 * Update a value cell
 */
function updateValueCell($pdo, $sheetId, $cellAddress, $rowNum, $colNum, $colLetter, $value) {
    // Remove from dependent cells if it was a formula
    $stmt = $pdo->prepare("DELETE FROM cells_dependent WHERE sheet_id = ? AND cell_address = ?");
    $stmt->execute([$sheetId, $cellAddress]);
    
    // Insert or update in independent cells
    $valueType = determineValueType($value);
    
    $stmt = $pdo->prepare("
        INSERT INTO cells_independent 
        (sheet_id, cell_address, row_num, col_num, col_letter, value, value_type, formatted_value)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ON DUPLICATE KEY UPDATE
            value = VALUES(value),
            value_type = VALUES(value_type),
            formatted_value = VALUES(formatted_value),
            updated_at = CURRENT_TIMESTAMP
    ");
    
    $stmt->execute([
        $sheetId, $cellAddress, $rowNum, $colNum, $colLetter,
        $value, $valueType, $value
    ]);
}

/**
 * Find ALL cells that depend on the given cell
 * This includes direct references and VLOOKUP formulas that use this cell
 */
function findAllDependentCells($pdo, $sheetId, $cellAddress) {
    $dependentCells = [];
    $processed = [];
    $toProcess = [$cellAddress];
    
    while (!empty($toProcess)) {
        $currentCell = array_shift($toProcess);
        
        if (in_array($currentCell, $processed)) {
            continue;
        }
        
        $processed[] = $currentCell;
        
        // Find cells with direct references
        $directDependents = findDirectDependents($pdo, $sheetId, $currentCell);
        
        // Find VLOOKUP formulas that might be affected
        $vlookupDependents = findVLOOKUPDependents($pdo, $sheetId, $currentCell);
        
        // Find MATCH formulas that might be affected
        $matchDependents = findMATCHDependents($pdo, $sheetId, $currentCell);
        
        // Combine all dependents
        $allDependents = array_unique(array_merge($directDependents, $vlookupDependents, $matchDependents));
        
        foreach ($allDependents as $dependent) {
            if (!in_array($dependent, $dependentCells) && $dependent !== $cellAddress) {
                $dependentCells[] = $dependent;
                $toProcess[] = $dependent;
            }
        }
    }
    
    // Sort dependents by their dependency depth to ensure correct calculation order
    return sortByDependencyDepth($pdo, $sheetId, $dependentCells);
}

/**
 * Find cells that directly reference the given cell in their formulas
 */
function findDirectDependents($pdo, $sheetId, $cellAddress) {
    $dependents = [];
    
    // Clean the cell address (remove $ signs)
    $cleanAddress = str_replace('$', '', $cellAddress);
    
    // Look for cells that have this cell in their dependencies JSON
    $stmt = $pdo->prepare("
        SELECT cell_address, formula, dependencies
        FROM cells_dependent 
        WHERE sheet_id = ? 
        AND (
            dependencies LIKE ? 
            OR dependencies LIKE ? 
            OR dependencies LIKE ?
            OR dependencies LIKE ?
            OR formula LIKE ?
            OR formula LIKE ?
            OR formula LIKE ?
            OR formula LIKE ?
        )
    ");
    
    $patterns = [
        '%"' . $cleanAddress . '"%',           // In JSON array
        '%["' . $cleanAddress . '"%',          // Start of JSON array
        '%,"' . $cleanAddress . '"%',          // Middle of JSON array
        '%"' . $cleanAddress . '"]%',          // End of JSON array
        '%' . $cleanAddress . '%',              // Direct reference in formula
        '%$' . $cleanAddress . '%',             // With $ sign
        '%' . $cleanAddress . ',%',            // Before comma
        '%' . $cleanAddress . ')%'             // Before closing paren
    ];
    
    $stmt->execute([$sheetId, ...$patterns]);
    
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        // Double-check that this cell actually references our target
        if (formulaReferencesCell($row['formula'], $cleanAddress)) {
            $dependents[] = $row['cell_address'];
        }
    }
    
    return $dependents;
}

/**
 * Find VLOOKUP formulas that might be affected by this cell change
 */
function findVLOOKUPDependents($pdo, $sheetId, $cellAddress) {
    $dependents = [];
    
    // Get all VLOOKUP formulas in the sheet
    $stmt = $pdo->prepare("
        SELECT cell_address, formula 
        FROM cells_dependent 
        WHERE sheet_id = ? 
        AND formula_type = 'VLOOKUP'
    ");
    $stmt->execute([$sheetId]);
    
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        // Check if this VLOOKUP references our cell
        if (vlookupReferencesCell($row['formula'], $cellAddress)) {
            $dependents[] = $row['cell_address'];
        }
    }
    
    return $dependents;
}

/**
 * Find MATCH formulas that might be affected
 */
function findMATCHDependents($pdo, $sheetId, $cellAddress) {
    $dependents = [];
    
    // Get all formulas containing MATCH
    $stmt = $pdo->prepare("
        SELECT cell_address, formula 
        FROM cells_dependent 
        WHERE sheet_id = ? 
        AND formula LIKE '%MATCH%'
    ");
    $stmt->execute([$sheetId]);
    
    while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
        if (formulaReferencesCell($row['formula'], $cellAddress)) {
            $dependents[] = $row['cell_address'];
        }
    }
    
    return $dependents;
}

/**
 * Check if a formula references a specific cell
 */
function formulaReferencesCell($formula, $cellAddress) {
    // Remove $ signs for comparison
    $cleanCell = str_replace('$', '', $cellAddress);
    $cleanFormula = str_replace('$', '', $formula);
    
    // Parse cell address
    if (!preg_match('/([A-Z]+)(\d+)/', $cleanCell, $cellMatch)) {
        return false;
    }
    
    $targetCol = $cellMatch[1];
    $targetRow = $cellMatch[2];
    
    // Check for direct cell reference
    if (preg_match('/\b' . $targetCol . $targetRow . '\b/', $cleanFormula)) {
        return true;
    }
    
    // Check if cell is within any range references in the formula
    preg_match_all('/([A-Z]+)(\d+):([A-Z]+)(\d+)/', $cleanFormula, $ranges, PREG_SET_ORDER);
    
    foreach ($ranges as $range) {
        $startCol = columnLetterToNumber($range[1]);
        $startRow = intval($range[2]);
        $endCol = columnLetterToNumber($range[3]);
        $endRow = intval($range[4]);
        
        $cellCol = columnLetterToNumber($targetCol);
        $cellRow = intval($targetRow);
        
        if ($cellCol >= $startCol && $cellCol <= $endCol && 
            $cellRow >= $startRow && $cellRow <= $endRow) {
            return true;
        }
    }
    
    return false;
}

/**
 * Check if a VLOOKUP formula references a specific cell
 */
function vlookupReferencesCell($formula, $cellAddress) {
    // First check if the lookup value references the cell
    if (preg_match('/VLOOKUP\s*\(\s*([^,]+),/i', $formula, $matches)) {
        $lookupValue = trim($matches[1]);
        
        // Clean addresses for comparison
        $cleanCell = str_replace('$', '', $cellAddress);
        $cleanLookup = str_replace('$', '', $lookupValue);
        
        if ($cleanLookup === $cleanCell) {
            return true;
        }
    }
    
    // Also check if the cell is in the table array
    return formulaReferencesCell($formula, $cellAddress);
}

/**
 * Sort cells by dependency depth to ensure correct calculation order
 */
function sortByDependencyDepth($pdo, $sheetId, $cells) {
    if (empty($cells)) {
        return [];
    }
    
    // Build dependency graph
    $graph = [];
    $depths = [];
    
    foreach ($cells as $cell) {
        $stmt = $pdo->prepare("
            SELECT dependencies 
            FROM cells_dependent 
            WHERE sheet_id = ? AND cell_address = ?
        ");
        $stmt->execute([$sheetId, $cell]);
        $row = $stmt->fetch(PDO::FETCH_ASSOC);
        
        if ($row && $row['dependencies']) {
            $deps = json_decode($row['dependencies'], true);
            if (is_array($deps)) {
                $graph[$cell] = array_map(function($d) {
                    return str_replace('$', '', $d);
                }, $deps);
            } else {
                $graph[$cell] = [];
            }
        } else {
            $graph[$cell] = [];
        }
    }
    
    // Calculate depth for each cell
    foreach ($cells as $cell) {
        $depths[$cell] = calculateDepth($cell, $graph, []);
    }
    
    // Sort by depth (cells with no dependencies first)
    usort($cells, function($a, $b) use ($depths) {
        return ($depths[$a] ?? 0) - ($depths[$b] ?? 0);
    });
    
    return $cells;
}

/**
 * Calculate dependency depth of a cell
 */
function calculateDepth($cell, $graph, $visited) {
    if (in_array($cell, $visited)) {
        return 0; // Circular reference
    }
    
    if (!isset($graph[$cell]) || empty($graph[$cell])) {
        return 0;
    }
    
    $visited[] = $cell;
    $maxDepth = 0;
    
    foreach ($graph[$cell] as $dep) {
        if (isset($graph[$dep])) {
            $depth = calculateDepth($dep, $graph, $visited);
            $maxDepth = max($maxDepth, $depth);
        }
    }
    
    return $maxDepth + 1;
}

/**
 * Helper function to convert column letter to number
 */
function columnLetterToNumber($letter) {
    $num = 0;
    for ($i = 0; $i < strlen($letter); $i++) {
        $num = $num * 26 + (ord($letter[$i]) - 64);
    }
    return $num;
}

/**
 * Extract formula type
 */
function extractFormulaType($formula) {
    if (preg_match('/^=([A-Z]+)\(/i', $formula, $matches)) {
        return strtoupper($matches[1]);
    }
    return 'CUSTOM';
}

/**
 * Extract dependencies from formula
 */
function extractDependencies($formula) {
    $dependencies = [];
    
    // Remove the = sign
    $formula = ltrim($formula, '=');
    
    // Match cell references (including with $)
    preg_match_all('/\$?[A-Z]+\$?\d+/i', $formula, $matches);
    
    if (!empty($matches[0])) {
        // Clean and deduplicate
        $dependencies = array_map(function($ref) {
            return str_replace('$', '', $ref);
        }, $matches[0]);
        
        $dependencies = array_unique($dependencies);
    }
    
    return array_values($dependencies);
}

/**
 * Determine value type
 */
function determineValueType($value) {
    if (is_null($value) || $value === '') return 'null';
    if (is_bool($value) || in_array(strtolower($value), ['true', 'false'])) return 'boolean';
    if (is_numeric($value)) return 'number';
    return 'string';
}