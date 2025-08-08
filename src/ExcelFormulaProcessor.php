<?php
/**
 * Excel Formula Processing Engine for Web Application
 * This handles all formula calculations just like Excel
 */

class ExcelFormulaProcessor {
    private $pdo;
    private $sheetId;
    private $cells = [];
    private $formulas = [];
    private $dependencies = [];
    private $calculationCache = [];
    
    public function __construct($pdo, $sheetId) {
        $this->pdo = $pdo;
        $this->sheetId = $sheetId;
        $this->loadSheetData();
    }
    
    /**
     * Load all sheet data into memory for processing
     */
    private function loadSheetData() {
        // Load all cells
        $stmt = $this->pdo->prepare("
            SELECT 
                cell_address, value, formula, formula_type, 
                dependencies, external_references
            FROM (
                SELECT cell_address, value, NULL as formula, NULL as formula_type, 
                       NULL as dependencies, NULL as external_references
                FROM cells_independent WHERE sheet_id = ?
                UNION ALL
                SELECT cell_address, calculated_value as value, formula, formula_type,
                       dependencies, external_references
                FROM cells_dependent WHERE sheet_id = ?
            ) cells
        ");
        $stmt->execute([$this->sheetId, $this->sheetId]);
        
        while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
            $this->cells[$row['cell_address']] = $row['value'];
            
            if ($row['formula']) {
                $this->formulas[$row['cell_address']] = $row['formula'];
                
                if ($row['dependencies']) {
                    $this->dependencies[$row['cell_address']] = json_decode($row['dependencies'], true);
                }
            }
        }
    }
    
    /**
     * Main formula evaluation function
     */
    public function evaluateFormula($formula, $cellAddress = null) {
        // Remove leading =
        $formula = ltrim($formula, '=');
        
        // Check for circular reference
        if ($cellAddress && $this->hasCircularReference($cellAddress, $formula)) {
            return ['value' => '#REF!', 'error' => 'Circular reference detected'];
        }
        
        try {
            // Parse and evaluate based on formula type
            if (preg_match('/^VLOOKUP\(/i', $formula)) {
                return $this->evaluateVLOOKUP($formula);
            } elseif (preg_match('/^MATCH\(/i', $formula)) {
                return $this->evaluateMATCH($formula);
            } elseif (preg_match('/^IFERROR\(/i', $formula)) {
                return $this->evaluateIFERROR($formula);
            } elseif (preg_match('/^SUM\(/i', $formula)) {
                return $this->evaluateSUM($formula);
            } elseif (preg_match('/^IF\(/i', $formula)) {
                return $this->evaluateIF($formula);
            } elseif (preg_match('/^CONCATENATE\(/i', $formula)) {
                return $this->evaluateCONCATENATE($formula);
            } else {
                // Handle arithmetic and cell references
                return $this->evaluateExpression($formula);
            }
        } catch (Exception $e) {
            return ['value' => '#ERROR!', 'error' => $e->getMessage()];
        }
    }
    
    /**
     * VLOOKUP Implementation
     * =VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
     */
    private function evaluateVLOOKUP($formula) {
        // Parse VLOOKUP parameters
        $params = $this->parseFormulaParams($formula, 'VLOOKUP');
        
        if (count($params) < 3) {
            return ['value' => '#VALUE!', 'error' => 'VLOOKUP requires at least 3 parameters'];
        }
        
        // Get lookup value
        $lookupValue = $this->resolveValue($params[0]);
        
        // Parse table range
        $tableRange = $this->parseRange($params[1]);
        
        // Get column index
        $colIndex = intval($this->resolveValue($params[2]));
        
        // Exact match flag (default false for approximate match)
        $exactMatch = isset($params[3]) && 
                     (strtoupper($params[3]) === 'FALSE' || $params[3] === '0');
        
        // Perform lookup
        $result = $this->performVLOOKUP($lookupValue, $tableRange, $colIndex, $exactMatch);
        
        return ['value' => $result];
    }
    
    /**
     * Perform actual VLOOKUP operation
     */
    private function performVLOOKUP($lookupValue, $tableRange, $colIndex, $exactMatch) {
        // Get table data
        $tableData = $this->getRangeData($tableRange);
        
        if (empty($tableData)) {
            return '#N/A';
        }
        
        // Search in first column
        foreach ($tableData as $row) {
            if (empty($row)) continue;
            
            $firstCol = reset($row);
            
            if ($exactMatch) {
                // Exact match
                if ($firstCol == $lookupValue) {
                    // Return value from specified column
                    $values = array_values($row);
                    return isset($values[$colIndex - 1]) ? $values[$colIndex - 1] : '#REF!';
                }
            } else {
                // Approximate match (for sorted data)
                if ($firstCol <= $lookupValue) {
                    $values = array_values($row);
                    $result = isset($values[$colIndex - 1]) ? $values[$colIndex - 1] : '#REF!';
                } else {
                    break;
                }
            }
        }
        
        return isset($result) ? $result : '#N/A';
    }
    
    /**
     * MATCH Implementation
     * =MATCH(lookup_value, lookup_array, [match_type])
     */
    private function evaluateMATCH($formula) {
        $params = $this->parseFormulaParams($formula, 'MATCH');
        
        if (count($params) < 2) {
            return ['value' => '#VALUE!', 'error' => 'MATCH requires at least 2 parameters'];
        }
        
        $lookupValue = $this->resolveValue($params[0]);
        $lookupRange = $this->parseRange($params[1]);
        $matchType = isset($params[2]) ? intval($params[2]) : 1;
        
        // Get range data
        $rangeData = $this->getRangeData($lookupRange);
        $values = [];
        
        // Flatten to single array
        foreach ($rangeData as $row) {
            foreach ($row as $value) {
                $values[] = $value;
            }
        }
        
        // Find match based on type
        if ($matchType == 0) {
            // Exact match
            $position = array_search($lookupValue, $values);
            return ['value' => $position !== false ? $position + 1 : '#N/A'];
        } elseif ($matchType == 1) {
            // Largest value less than or equal to lookup_value
            $position = 0;
            foreach ($values as $i => $value) {
                if ($value <= $lookupValue) {
                    $position = $i + 1;
                } else {
                    break;
                }
            }
            return ['value' => $position > 0 ? $position : '#N/A'];
        } else {
            // Smallest value greater than or equal to lookup_value
            foreach ($values as $i => $value) {
                if ($value >= $lookupValue) {
                    return ['value' => $i + 1];
                }
            }
            return ['value' => '#N/A'];
        }
    }
    
    /**
     * IFERROR Implementation
     * =IFERROR(value, value_if_error)
     */
    private function evaluateIFERROR($formula) {
        $params = $this->parseFormulaParams($formula, 'IFERROR');
        
        if (count($params) < 2) {
            return ['value' => '#VALUE!', 'error' => 'IFERROR requires 2 parameters'];
        }
        
        // Try to evaluate first parameter
        $result = $this->evaluateFormula($params[0]);
        
        // Check if it's an error
        if (is_array($result) && isset($result['error'])) {
            // Return second parameter
            return $this->evaluateFormula($params[1]);
        } elseif (is_string($result['value']) && strpos($result['value'], '#') === 0) {
            // Excel error value
            return $this->evaluateFormula($params[1]);
        }
        
        return $result;
    }
    
    /**
     * SUM Implementation
     * =SUM(number1, [number2], ...)
     */
    private function evaluateSUM($formula) {
        $params = $this->parseFormulaParams($formula, 'SUM');
        $sum = 0;
        
        foreach ($params as $param) {
            // Check if it's a range
            if (strpos($param, ':') !== false) {
                $rangeData = $this->getRangeData($this->parseRange($param));
                foreach ($rangeData as $row) {
                    foreach ($row as $value) {
                        if (is_numeric($value)) {
                            $sum += floatval($value);
                        }
                    }
                }
            } else {
                // Single cell or value
                $value = $this->resolveValue($param);
                if (is_numeric($value)) {
                    $sum += floatval($value);
                }
            }
        }
        
        return ['value' => $sum];
    }
    
    /**
     * IF Implementation
     * =IF(logical_test, value_if_true, value_if_false)
     */
    private function evaluateIF($formula) {
        $params = $this->parseFormulaParams($formula, 'IF');
        
        if (count($params) < 2) {
            return ['value' => '#VALUE!', 'error' => 'IF requires at least 2 parameters'];
        }
        
        // Evaluate condition
        $condition = $this->evaluateCondition($params[0]);
        
        // Return appropriate value
        if ($condition) {
            return $this->evaluateFormula($params[1]);
        } else {
            return isset($params[2]) ? $this->evaluateFormula($params[2]) : ['value' => 'FALSE'];
        }
    }
    
    /**
     * Evaluate arithmetic expressions and cell references
     */
    private function evaluateExpression($expression) {
        // Replace cell references with values
        $expression = preg_replace_callback(
            '/\b([A-Z]+\d+)\b/',
            function($matches) {
                return $this->getCellValue($matches[1]);
            },
            $expression
        );
        
        // Handle basic arithmetic
        // WARNING: eval() is dangerous - in production use a proper expression parser
        // For now, we'll do basic replacements
        try {
            // Simple arithmetic evaluation (safe for basic operations)
            $result = $this->safeEvaluateExpression($expression);
            return ['value' => $result];
        } catch (Exception $e) {
            return ['value' => '#VALUE!', 'error' => 'Invalid expression'];
        }
    }
    
    /**
     * Safe expression evaluation (without eval)
     */
    private function safeEvaluateExpression($expression) {
        // This is a simplified version - in production, use a proper expression parser
        // For now, handle basic arithmetic
        
        // Remove spaces
        $expression = str_replace(' ', '', $expression);
        
        // If it's just a number, return it
        if (is_numeric($expression)) {
            return $expression;
        }
        
        // Handle concatenation (&)
        if (strpos($expression, '&') !== false) {
            $parts = explode('&', $expression);
            $result = '';
            foreach ($parts as $part) {
                $result .= trim($part, '"');
            }
            return $result;
        }
        
        // For complex arithmetic, you'd need a proper parser
        // This is a placeholder
        return $expression;
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
                elseif ($char == ',' && $depth == 0) {
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
     * Parse range reference (e.g., A1:B10 or Sheet1!A1:B10)
     */
    private function parseRange($rangeStr) {
        $range = [];
        
        // Check for sheet reference
        if (strpos($rangeStr, '!') !== false) {
            list($sheetName, $rangeStr) = explode('!', $rangeStr);
            $range['sheet'] = trim($sheetName, "'");
        } else {
            $range['sheet'] = null;
        }
        
        // Parse range
        if (strpos($rangeStr, ':') !== false) {
            list($start, $end) = explode(':', $rangeStr);
            $range['start'] = $this->parseCellAddress($start);
            $range['end'] = $this->parseCellAddress($end);
        } else {
            $range['start'] = $this->parseCellAddress($rangeStr);
            $range['end'] = $range['start'];
        }
        
        return $range;
    }
    
    /**
     * Parse cell address into row and column
     */
    private function parseCellAddress($address) {
        $address = str_replace('$', '', $address);
        preg_match('/([A-Z]+)(\d+)/i', $address, $matches);
        
        if (!$matches) {
            return null;
        }
        
        return [
            'col' => $this->columnToNumber($matches[1]),
            'row' => intval($matches[2]),
            'address' => $address
        ];
    }
    
    /**
     * Get data from a range
     */
    private function getRangeData($range) {
        $data = [];
        
        // Determine sheet
        $sheetId = $this->sheetId;
        if ($range['sheet']) {
            // Get sheet ID from name
            $stmt = $this->pdo->prepare("
                SELECT id FROM sheets 
                WHERE sheet_name = ? 
                AND workbook_id = (
                    SELECT workbook_id FROM sheets WHERE id = ?
                )
            ");
            $stmt->execute([$range['sheet'], $this->sheetId]);
            $sheetId = $stmt->fetchColumn();
        }
        
        // Get cells in range
        for ($row = $range['start']['row']; $row <= $range['end']['row']; $row++) {
            $rowData = [];
            for ($col = $range['start']['col']; $col <= $range['end']['col']; $col++) {
                $address = $this->numberToColumn($col) . $row;
                $rowData[$address] = $this->getCellValueFromDB($address, $sheetId);
            }
            $data[] = $rowData;
        }
        
        return $data;
    }
    
    /**
     * Get cell value
     */
    private function getCellValue($address) {
        if (isset($this->cells[$address])) {
            return $this->cells[$address];
        }
        return 0;
    }
    
    /**
     * Get cell value from database
     */
    private function getCellValueFromDB($address, $sheetId) {
        $stmt = $this->pdo->prepare("
            SELECT value FROM cells_independent 
            WHERE sheet_id = ? AND cell_address = ?
            UNION
            SELECT calculated_value FROM cells_dependent 
            WHERE sheet_id = ? AND cell_address = ?
            LIMIT 1
        ");
        $stmt->execute([$sheetId, $address, $sheetId, $address]);
        
        $value = $stmt->fetchColumn();
        return $value !== false ? $value : '';
    }
    
    /**
     * Resolve value (cell reference or literal)
     */
    private function resolveValue($value) {
        $value = trim($value);
        
        // Remove quotes for string literals
        if (preg_match('/^"(.*)"$/', $value, $matches)) {
            return $matches[1];
        }
        
        // Check if it's a cell reference
        if (preg_match('/^[A-Z]+\d+$/i', $value)) {
            return $this->getCellValue($value);
        }
        
        // Check if it's a number
        if (is_numeric($value)) {
            return $value;
        }
        
        // Boolean values
        if (strtoupper($value) === 'TRUE') return true;
        if (strtoupper($value) === 'FALSE') return false;
        
        return $value;
    }
    
    /**
     * Evaluate condition for IF statements
     */
    private function evaluateCondition($condition) {
        // Parse comparison operators
        if (preg_match('/(.+?)(>=|<=|<>|>|<|=)(.+)/', $condition, $matches)) {
            $left = $this->resolveValue(trim($matches[1]));
            $operator = $matches[2];
            $right = $this->resolveValue(trim($matches[3]));
            
            switch ($operator) {
                case '=':
                    return $left == $right;
                case '<>':
                    return $left != $right;
                case '>':
                    return $left > $right;
                case '<':
                    return $left < $right;
                case '>=':
                    return $left >= $right;
                case '<=':
                    return $left <= $right;
            }
        }
        
        // Simple boolean
        $value = $this->resolveValue($condition);
        return $value && $value !== '0' && $value !== 'FALSE';
    }
    
    /**
     * Check for circular references
     */
    private function hasCircularReference($cellAddress, $formula, $visited = []) {
        if (in_array($cellAddress, $visited)) {
            return true;
        }
        
        $visited[] = $cellAddress;
        
        // Extract cell references from formula
        preg_match_all('/\b([A-Z]+\d+)\b/', $formula, $matches);
        
        foreach ($matches[1] as $ref) {
            if (isset($this->formulas[$ref])) {
                if ($this->hasCircularReference($ref, $this->formulas[$ref], $visited)) {
                    return true;
                }
            }
        }
        
        return false;
    }
    
    /**
     * Column letter to number conversion
     */
    private function columnToNumber($col) {
        $num = 0;
        for ($i = 0; $i < strlen($col); $i++) {
            $num = $num * 26 + (ord($col[$i]) - 64);
        }
        return $num;
    }
    
    /**
     * Number to column letter conversion
     */
    private function numberToColumn($num) {
        $col = '';
        while ($num > 0) {
            $num--;
            $col = chr(65 + ($num % 26)) . $col;
            $num = intval($num / 26);
        }
        return $col;
    }
    
    /**
     * Update cell and recalculate dependents
     */
    public function updateCell($cellAddress, $value, $isFormula = false) {
        try {
            $this->pdo->beginTransaction();
            
            if ($isFormula) {
                // Evaluate formula
                $result = $this->evaluateFormula($value, $cellAddress);
                $calculatedValue = is_array($result) ? $result['value'] : $result;
                
                // Extract formula type
                preg_match('/^=([A-Z]+)\(/', $value, $matches);
                $formulaType = isset($matches[1]) ? $matches[1] : 'CUSTOM';
                
                // Extract dependencies
                preg_match_all('/\b([A-Z]+\d+)\b/', $value, $matches);
                $dependencies = array_unique($matches[1]);
                
                // Check if cell exists
                $stmt = $this->pdo->prepare("
                    SELECT id FROM cells_dependent 
                    WHERE sheet_id = ? AND cell_address = ?
                ");
                $stmt->execute([$this->sheetId, $cellAddress]);
                
                if ($stmt->fetchColumn()) {
                    // Update existing
                    $stmt = $this->pdo->prepare("
                        UPDATE cells_dependent 
                        SET formula = ?, formula_type = ?, calculated_value = ?,
                            dependencies = ?, has_error = ?, error_message = ?
                        WHERE sheet_id = ? AND cell_address = ?
                    ");
                } else {
                    // Insert new
                    $stmt = $this->pdo->prepare("
                        INSERT INTO cells_dependent 
                        (formula, formula_type, calculated_value, dependencies, 
                         has_error, error_message, sheet_id, cell_address,
                         row_num, col_num, col_letter)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ");
                }
                
                $hasError = isset($result['error']);
                $errorMsg = $hasError ? $result['error'] : null;
                
                // Parse cell address
                preg_match('/([A-Z]+)(\d+)/', $cellAddress, $addrMatch);
                $colLetter = $addrMatch[1];
                $rowNum = $addrMatch[2];
                $colNum = $this->columnToNumber($colLetter);
                
                if ($stmt->columnCount() > 8) {
                    // Insert with position data
                    $stmt->execute([
                        $value, $formulaType, $calculatedValue,
                        json_encode($dependencies), $hasError, $errorMsg,
                        $this->sheetId, $cellAddress, $rowNum, $colNum, $colLetter
                    ]);
                } else {
                    // Update
                    $stmt->execute([
                        $value, $formulaType, $calculatedValue,
                        json_encode($dependencies), $hasError, $errorMsg,
                        $this->sheetId, $cellAddress
                    ]);
                }
                
                // Delete from independent if exists
                $stmt = $this->pdo->prepare("
                    DELETE FROM cells_independent 
                    WHERE sheet_id = ? AND cell_address = ?
                ");
                $stmt->execute([$this->sheetId, $cellAddress]);
                
            } else {
                // Regular value
                $valueType = $this->determineValueType($value);
                
                // Check if cell exists
                $stmt = $this->pdo->prepare("
                    SELECT id FROM cells_independent 
                    WHERE sheet_id = ? AND cell_address = ?
                ");
                $stmt->execute([$this->sheetId, $cellAddress]);
                
                if ($stmt->fetchColumn()) {
                    // Update existing
                    $stmt = $this->pdo->prepare("
                        UPDATE cells_independent 
                        SET value = ?, value_type = ?
                        WHERE sheet_id = ? AND cell_address = ?
                    ");
                    $stmt->execute([$value, $valueType, $this->sheetId, $cellAddress]);
                } else {
                    // Insert new
                    preg_match('/([A-Z]+)(\d+)/', $cellAddress, $addrMatch);
                    $colLetter = $addrMatch[1];
                    $rowNum = $addrMatch[2];
                    $colNum = $this->columnToNumber($colLetter);
                    
                    $stmt = $this->pdo->prepare("
                        INSERT INTO cells_independent 
                        (sheet_id, cell_address, row_num, col_num, col_letter, 
                         value, value_type)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ");
                    $stmt->execute([
                        $this->sheetId, $cellAddress, $rowNum, $colNum, $colLetter,
                        $value, $valueType
                    ]);
                }
                
                // Delete from dependent if exists
                $stmt = $this->pdo->prepare("
                    DELETE FROM cells_dependent 
                    WHERE sheet_id = ? AND cell_address = ?
                ");
                $stmt->execute([$this->sheetId, $cellAddress]);
            }
            
            $this->pdo->commit();
            
            // Update cache
            $this->cells[$cellAddress] = $isFormula ? $calculatedValue : $value;
            
            // Get affected cells (dependents)
            $affectedCells = $this->getDependentCells($cellAddress);
            
            return [
                'success' => true,
                'cell' => [
                    'cell_address' => $cellAddress,
                    'value' => $isFormula ? $calculatedValue : $value,
                    'formula' => $isFormula ? $value : null,
                    'formatted_value' => $isFormula ? $calculatedValue : $value
                ],
                'affected_cells' => $affectedCells
            ];
            
        } catch (Exception $e) {
            $this->pdo->rollBack();
            return [
                'success' => false,
                'error' => $e->getMessage()
            ];
        }
    }
    
    /**
     * Get cells that depend on a given cell
     */
    private function getDependentCells($cellAddress) {
        $dependents = [];
        
        $stmt = $this->pdo->prepare("
            SELECT cell_address, formula 
            FROM cells_dependent 
            WHERE sheet_id = ? 
            AND dependencies LIKE ?
        ");
        $stmt->execute([$this->sheetId, '%"' . $cellAddress . '"%']);
        
        while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
            $dependents[] = $row['cell_address'];
        }
        
        return $dependents;
    }
    
    /**
     * Determine value type
     */
    private function determineValueType($value) {
        if (is_null($value) || $value === '') return 'null';
        if (is_bool($value)) return 'boolean';
        if (is_numeric($value)) return 'number';
        if (strtotime($value) !== false) return 'date';
        return 'string';
    }
}
