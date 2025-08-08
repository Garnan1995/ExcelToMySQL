<?php
/**
 * Enhanced Formula Recalculation Engine
 * Properly evaluates Excel formulas including VLOOKUP, MATCH, etc.
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

class FormulaEvaluator {
    private $pdo;
    private $sheetId;
    private $cache = [];
    private $evaluationStack = [];
    
    public function __construct($pdo, $sheetId) {
        $this->pdo = $pdo;
        $this->sheetId = $sheetId;
    }
    
    /**
     * Main evaluation entry point
     */
    public function evaluate($cellAddress, $formula) {
        // Check for circular reference
        if (in_array($cellAddress, $this->evaluationStack)) {
            return '#REF!';
        }
        
        $this->evaluationStack[] = $cellAddress;
        
        try {
            $result = $this->evaluateFormula($formula);
            array_pop($this->evaluationStack);
            return $result;
        } catch (Exception $e) {
            array_pop($this->evaluationStack);
            return '#ERROR!';
        }
    }
    
    /**
     * Evaluate a formula
     */
    private function evaluateFormula($formula) {
        // Remove leading =
        $formula = ltrim($formula, '=');
        
        // Handle different formula types
        if (preg_match('/^VLOOKUP\s*\(/i', $formula)) {
            return $this->evaluateVLOOKUP($formula);
        }
        
        if (preg_match('/^MATCH\s*\(/i', $formula)) {
            return $this->evaluateMATCH($formula);
        }
        
        if (preg_match('/^IFERROR\s*\(/i', $formula)) {
            return $this->evaluateIFERROR($formula);
        }
        
        if (preg_match('/^SUM\s*\(/i', $formula)) {
            return $this->evaluateSUM($formula);
        }
        
        if (preg_match('/^IF\s*\(/i', $formula)) {
            return $this->evaluateIF($formula);
        }
        
        // Handle simple cell reference
        if (preg_match('/^([A-Z]+\d+)$/i', $formula)) {
            return $this->getCellValue($formula);
        }
        
        // Handle arithmetic expressions
        return $this->evaluateExpression($formula);
    }
    
    /**
     * VLOOKUP Implementation
     */
    private function evaluateVLOOKUP($formula) {
        $params = $this->parseFormulaParams($formula, 'VLOOKUP');
        
        if (count($params) < 3) {
            return '#VALUE!';
        }
        
        // Get lookup value
        $lookupValue = $this->resolveValue($params[0]);
        
        // Parse table range
        $tableRange = $this->parseRange($params[1]);
        
        // Get column index
        $colIndex = intval($this->resolveValue($params[2]));
        
        // Exact match flag (default FALSE for approximate match)
        $exactMatch = isset($params[3]) && 
                     (strtoupper(trim($params[3])) === 'FALSE' || 
                      trim($params[3]) === '0');
        
        // Perform the lookup
        return $this->performVLOOKUP($lookupValue, $tableRange, $colIndex, $exactMatch);
    }
    
    /**
     * Perform actual VLOOKUP operation
     */
    private function performVLOOKUP($lookupValue, $tableRange, $colIndex, $exactMatch) {
        // Get the sheet ID for the table range
        $targetSheetId = $this->sheetId;
        
        if ($tableRange['sheet']) {
            $stmt = $this->pdo->prepare("
                SELECT id FROM sheets 
                WHERE sheet_name = ? 
                AND workbook_id = (
                    SELECT workbook_id FROM sheets WHERE id = ?
                )
            ");
            $stmt->execute([$tableRange['sheet'], $this->sheetId]);
            $targetSheetId = $stmt->fetchColumn();
            
            if (!$targetSheetId) {
                return '#REF!';
            }
        }
        
        // Build the query to get all rows in the range
        $startCol = $tableRange['startCol'];
        $endCol = $tableRange['endCol'];
        $startRow = $tableRange['startRow'];
        $endRow = $tableRange['endRow'];
        
        // Get all values in the first column of the range
        $sql = "
            SELECT 
                ci.row_num,
                ci.value as lookup_col_value
            FROM cells_independent ci
            WHERE ci.sheet_id = ?
                AND ci.col_letter = ?
                AND ci.row_num BETWEEN ? AND ?
                AND ci.value = ?
            
            UNION
            
            SELECT 
                cd.row_num,
                cd.calculated_value as lookup_col_value
            FROM cells_dependent cd
            WHERE cd.sheet_id = ?
                AND cd.col_letter = ?
                AND cd.row_num BETWEEN ? AND ?
                AND cd.calculated_value = ?
            
            ORDER BY row_num
            LIMIT 1
        ";
        
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute([
            $targetSheetId, $startCol, $startRow, $endRow, $lookupValue,
            $targetSheetId, $startCol, $startRow, $endRow, $lookupValue
        ]);
        
        $matchRow = $stmt->fetch(PDO::FETCH_ASSOC);
        
        if (!$matchRow) {
            // If not exact match and approximate match is allowed
            if (!$exactMatch) {
                // Find the largest value less than or equal to lookup value
                $sql = "
                    SELECT row_num, value as lookup_col_value
                    FROM (
                        SELECT row_num, value
                        FROM cells_independent
                        WHERE sheet_id = ? AND col_letter = ? 
                        AND row_num BETWEEN ? AND ?
                        AND CAST(value AS DECIMAL(10,2)) <= CAST(? AS DECIMAL(10,2))
                        
                        UNION
                        
                        SELECT row_num, calculated_value as value
                        FROM cells_dependent
                        WHERE sheet_id = ? AND col_letter = ?
                        AND row_num BETWEEN ? AND ?
                        AND CAST(calculated_value AS DECIMAL(10,2)) <= CAST(? AS DECIMAL(10,2))
                    ) t
                    ORDER BY CAST(lookup_col_value AS DECIMAL(10,2)) DESC
                    LIMIT 1
                ";
                
                $stmt = $this->pdo->prepare($sql);
                $stmt->execute([
                    $targetSheetId, $startCol, $startRow, $endRow, $lookupValue,
                    $targetSheetId, $startCol, $startRow, $endRow, $lookupValue
                ]);
                
                $matchRow = $stmt->fetch(PDO::FETCH_ASSOC);
            }
            
            if (!$matchRow) {
                return '#N/A';
            }
        }
        
        // Now get the value from the target column
        $targetCol = $this->columnOffsetFrom($startCol, $colIndex - 1);
        $targetRow = $matchRow['row_num'];
        
        $sql = "
            SELECT value FROM cells_independent 
            WHERE sheet_id = ? AND col_letter = ? AND row_num = ?
            
            UNION
            
            SELECT calculated_value as value FROM cells_dependent
            WHERE sheet_id = ? AND col_letter = ? AND row_num = ?
            
            LIMIT 1
        ";
        
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute([
            $targetSheetId, $targetCol, $targetRow,
            $targetSheetId, $targetCol, $targetRow
        ]);
        
        $result = $stmt->fetchColumn();
        
        return $result !== false ? $result : '#N/A';
    }
    
    /**
     * MATCH Implementation
     */
    private function evaluateMATCH($formula) {
        $params = $this->parseFormulaParams($formula, 'MATCH');
        
        if (count($params) < 2) {
            return '#VALUE!';
        }
        
        $lookupValue = $this->resolveValue($params[0]);
        $lookupRange = $this->parseRange($params[1]);
        $matchType = isset($params[2]) ? intval($params[2]) : 1;
        
        // Get the sheet ID
        $targetSheetId = $this->sheetId;
        if ($lookupRange['sheet']) {
            $stmt = $this->pdo->prepare("
                SELECT id FROM sheets WHERE sheet_name = ? 
                AND workbook_id = (SELECT workbook_id FROM sheets WHERE id = ?)
            ");
            $stmt->execute([$lookupRange['sheet'], $this->sheetId]);
            $targetSheetId = $stmt->fetchColumn();
        }
        
        // Get all values in range
        $sql = "
            SELECT 
                cell_address,
                CASE 
                    WHEN ci.value IS NOT NULL THEN ci.value
                    ELSE cd.calculated_value
                END as value,
                ROW_NUMBER() OVER (ORDER BY 
                    COALESCE(ci.row_num, cd.row_num), 
                    COALESCE(ci.col_num, cd.col_num)
                ) as position
            FROM (
                SELECT DISTINCT row_num, col_num 
                FROM (
                    SELECT row_num, col_num FROM cells_independent 
                    WHERE sheet_id = ? 
                    AND col_letter BETWEEN ? AND ?
                    AND row_num BETWEEN ? AND ?
                    UNION
                    SELECT row_num, col_num FROM cells_dependent
                    WHERE sheet_id = ?
                    AND col_letter BETWEEN ? AND ?
                    AND row_num BETWEEN ? AND ?
                ) t
            ) cells
            LEFT JOIN cells_independent ci ON 
                ci.sheet_id = ? AND ci.row_num = cells.row_num AND ci.col_num = cells.col_num
            LEFT JOIN cells_dependent cd ON
                cd.sheet_id = ? AND cd.row_num = cells.row_num AND cd.col_num = cells.col_num
        ";
        
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute([
            $targetSheetId, 
            $lookupRange['startCol'], $lookupRange['endCol'],
            $lookupRange['startRow'], $lookupRange['endRow'],
            $targetSheetId,
            $lookupRange['startCol'], $lookupRange['endCol'],
            $lookupRange['startRow'], $lookupRange['endRow'],
            $targetSheetId,
            $targetSheetId
        ]);
        
        $results = $stmt->fetchAll(PDO::FETCH_ASSOC);
        
        // Find match based on type
        if ($matchType == 0) {
            // Exact match
            foreach ($results as $row) {
                if ($row['value'] == $lookupValue) {
                    return $row['position'];
                }
            }
        } elseif ($matchType == 1) {
            // Largest value less than or equal
            $lastPosition = 0;
            foreach ($results as $row) {
                if ($row['value'] <= $lookupValue) {
                    $lastPosition = $row['position'];
                } else {
                    break;
                }
            }
            return $lastPosition > 0 ? $lastPosition : '#N/A';
        } else {
            // Smallest value greater than or equal
            foreach ($results as $row) {
                if ($row['value'] >= $lookupValue) {
                    return $row['position'];
                }
            }
        }
        
        return '#N/A';
    }
    
    /**
     * IFERROR Implementation
     */
    private function evaluateIFERROR($formula) {
        $params = $this->parseFormulaParams($formula, 'IFERROR');
        
        if (count($params) < 2) {
            return '#VALUE!';
        }
        
        // Try to evaluate the first parameter
        $result = $this->evaluateFormula($params[0]);
        
        // If it's an error, return the second parameter
        if (is_string($result) && strpos($result, '#') === 0) {
            return $this->evaluateFormula($params[1]);
        }
        
        return $result;
    }
    
    /**
     * Parse formula parameters
     */
    private function parseFormulaParams($formula, $functionName) {
        // Extract content between parentheses
        $pattern = '/' . $functionName . '\s*\((.*)\)$/i';
        if (!preg_match($pattern, $formula, $matches)) {
            return [];
        }
        
        $content = $matches[1];
        $params = [];
        $current = '';
        $depth = 0;
        $inQuotes = false;
        
        for ($i = 0; $i < strlen($content); $i++) {
            $char = $content[$i];
            
            if ($char == '"' && ($i == 0 || $content[$i-1] != '\\')) {
                $inQuotes = !$inQuotes;
            }
            
            if (!$inQuotes) {
                if ($char == '(') $depth++;
                elseif ($char == ')') $depth--;
                elseif ($char == ';' && $depth == 0) {
                    $params[] = trim($current);
                    $current = '';
                    continue;
                }
            }
            
            $current .= $char;
        }
        
        if ($current) {
            $params[] = trim($current);
        }
        
        return $params;
    }
    
    /**
     * Parse range reference
     */
    private function parseRange($rangeStr) {
        $range = [];
        
        // Check for sheet reference
        if (preg_match("/^'([^']+)'!/", $rangeStr, $matches)) {
            $range['sheet'] = $matches[1];
            $rangeStr = substr($rangeStr, strlen($matches[0]));
        } else {
            $range['sheet'] = null;
        }
        
        // Remove $ signs
        $rangeStr = str_replace('$', '', $rangeStr);
        
        // Parse range
        if (preg_match('/([A-Z]+)(\d+):([A-Z]+)(\d+)/', $rangeStr, $matches)) {
            $range['startCol'] = $matches[1];
            $range['startRow'] = intval($matches[2]);
            $range['endCol'] = $matches[3];
            $range['endRow'] = intval($matches[4]);
        } else {
            // Single cell
            preg_match('/([A-Z]+)(\d+)/', $rangeStr, $matches);
            $range['startCol'] = $matches[1] ?? 'A';
            $range['startRow'] = intval($matches[2] ?? 1);
            $range['endCol'] = $range['startCol'];
            $range['endRow'] = $range['startRow'];
        }
        
        return $range;
    }
    
    /**
     * Get cell value
     */
    private function getCellValue($cellAddress) {
        // Check cache first
        if (isset($this->cache[$cellAddress])) {
            return $this->cache[$cellAddress];
        }
        
        // Remove $ signs
        $cellAddress = str_replace('$', '', $cellAddress);
        
        $stmt = $this->pdo->prepare("
            SELECT value FROM cells_independent 
            WHERE sheet_id = ? AND cell_address = ?
            UNION
            SELECT calculated_value FROM cells_dependent 
            WHERE sheet_id = ? AND cell_address = ?
            LIMIT 1
        ");
        $stmt->execute([$this->sheetId, $cellAddress, $this->sheetId, $cellAddress]);
        
        $value = $stmt->fetchColumn();
        $value = $value !== false ? $value : '';
        
        // Cache the value
        $this->cache[$cellAddress] = $value;
        
        return $value;
    }
    
    /**
     * Resolve a value (could be literal or cell reference)
     */
    private function resolveValue($value) {
        $value = trim($value);
        
        // Remove quotes for string literals
        if (preg_match('/^"(.*)"$/', $value, $matches)) {
            return $matches[1];
        }
        
        // Check if it's a cell reference
        if (preg_match('/^[$]?[A-Z]+[$]?\d+$/i', $value)) {
            return $this->getCellValue($value);
        }
        
        // Check for nested function
        if (preg_match('/^[A-Z]+\(/i', $value)) {
            return $this->evaluateFormula($value);
        }
        
        // Return as literal
        return $value;
    }
    
    /**
     * Calculate column offset
     */
    private function columnOffsetFrom($startCol, $offset) {
        $num = 0;
        for ($i = 0; $i < strlen($startCol); $i++) {
            $num = $num * 26 + (ord($startCol[$i]) - 64);
        }
        
        $num += $offset;
        
        $col = '';
        while ($num > 0) {
            $num--;
            $col = chr(65 + ($num % 26)) . $col;
            $num = intval($num / 26);
        }
        
        return $col;
    }
    
    /**
     * Evaluate simple expressions
     */
    private function evaluateExpression($expression) {
        // Replace cell references with values
        $expression = preg_replace_callback(
            '/\b([A-Z]+\d+)\b/i',
            function($matches) {
                return $this->getCellValue($matches[1]);
            },
            $expression
        );
        
        // Handle basic arithmetic (simplified)
        // In production, use a proper expression parser
        if (is_numeric($expression)) {
            return $expression;
        }
        
        return $expression;
    }
}

// ============================================
// Main Execution
// ============================================

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
    
    $evaluator = new FormulaEvaluator($pdo, $sheetId);
    $results = [];
    
    // Recalculate each cell
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
            // Evaluate the formula
            $calculatedValue = $evaluator->evaluate($cellAddress, $cellData['formula']);
            
            // Update the calculated value
            $updateStmt = $pdo->prepare("
                UPDATE cells_dependent 
                SET calculated_value = ?, 
                    formatted_value = ?,
                    has_error = ?,
                    error_message = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE sheet_id = ? AND cell_address = ?
            ");
            
            $hasError = is_string($calculatedValue) && strpos($calculatedValue, '#') === 0;
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
        'error' => $e->getMessage(),
        'trace' => $e->getTraceAsString()
    ]);
}