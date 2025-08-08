<?php
/**
 * Excel to MySQL Converter with Complete Cell Metadata Preservation
 * Requires: composer require phpoffice/phpspreadsheet
 */

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Calculation\Calculation;

class ExcelToMySQLConverter {
    private $pdo;
    private $spreadsheet;
    private $workbookId;
    
    public function __construct($host, $dbname, $username, $password) {
        try {
            $this->pdo = new PDO("mysql:host=$host;dbname=$dbname;charset=utf8mb4", 
                                 $username, $password);
            $this->pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
            $this->createDatabaseSchema();
        } catch(PDOException $e) {
            die("Connection failed: " . $e->getMessage());
        }
    }
    
    /**
     * Create the database schema for storing Excel data
     */
    private function createDatabaseSchema() {
        // Main workbook table
        $this->pdo->exec("
            CREATE TABLE IF NOT EXISTS workbooks (
                id INT AUTO_INCREMENT PRIMARY KEY,
                filename VARCHAR(255) NOT NULL,
                original_path TEXT,
                import_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                last_modified TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                metadata JSON,
                INDEX idx_filename (filename)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        ");
        
        // Sheets table
        $this->pdo->exec("
            CREATE TABLE IF NOT EXISTS sheets (
                id INT AUTO_INCREMENT PRIMARY KEY,
                workbook_id INT NOT NULL,
                sheet_name VARCHAR(255) NOT NULL,
                sheet_index INT NOT NULL,
                is_active BOOLEAN DEFAULT TRUE,
                row_count INT DEFAULT 0,
                column_count INT DEFAULT 0,
                metadata JSON,
                FOREIGN KEY (workbook_id) REFERENCES workbooks(id) ON DELETE CASCADE,
                INDEX idx_workbook_sheet (workbook_id, sheet_name)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        ");
        
        // Independent cells (no formulas)
        $this->pdo->exec("
            CREATE TABLE IF NOT EXISTS cells_independent (
                id INT AUTO_INCREMENT PRIMARY KEY,
                sheet_id INT NOT NULL,
                cell_address VARCHAR(20) NOT NULL,
                row_num INT NOT NULL,
                col_num INT NOT NULL,
                col_letter VARCHAR(5) NOT NULL,
                value TEXT,
                value_type ENUM('string', 'number', 'boolean', 'date', 'null') DEFAULT 'null',
                formatted_value TEXT,
                number_format VARCHAR(255),
                hyperlink TEXT,
                comment TEXT,
                style JSON,
                metadata JSON,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                FOREIGN KEY (sheet_id) REFERENCES sheets(id) ON DELETE CASCADE,
                UNIQUE KEY unique_cell (sheet_id, cell_address),
                INDEX idx_sheet_row_col (sheet_id, row_num, col_num),
                INDEX idx_value_type (value_type)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        ");
        
        // Dependent cells (with formulas)
        $this->pdo->exec("
            CREATE TABLE IF NOT EXISTS cells_dependent (
                id INT AUTO_INCREMENT PRIMARY KEY,
                sheet_id INT NOT NULL,
                cell_address VARCHAR(20) NOT NULL,
                row_num INT NOT NULL,
                col_num INT NOT NULL,
                col_letter VARCHAR(5) NOT NULL,
                formula TEXT NOT NULL,
                formula_type VARCHAR(50),
                calculated_value TEXT,
                value_type ENUM('string', 'number', 'boolean', 'date', 'error', 'null') DEFAULT 'null',
                formatted_value TEXT,
                number_format VARCHAR(255),
                dependencies JSON COMMENT 'Array of cell references this formula depends on',
                external_references JSON COMMENT 'References to other sheets/workbooks',
                is_array_formula BOOLEAN DEFAULT FALSE,
                has_error BOOLEAN DEFAULT FALSE,
                error_message TEXT,
                hyperlink TEXT,
                comment TEXT,
                style JSON,
                metadata JSON,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                FOREIGN KEY (sheet_id) REFERENCES sheets(id) ON DELETE CASCADE,
                UNIQUE KEY unique_cell (sheet_id, cell_address),
                INDEX idx_sheet_row_col (sheet_id, row_num, col_num),
                INDEX idx_formula_type (formula_type),
                INDEX idx_has_error (has_error)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        ");
        
        // Named ranges table
        $this->pdo->exec("
            CREATE TABLE IF NOT EXISTS named_ranges (
                id INT AUTO_INCREMENT PRIMARY KEY,
                workbook_id INT NOT NULL,
                sheet_id INT,
                name VARCHAR(255) NOT NULL,
                range_address TEXT NOT NULL,
                scope ENUM('workbook', 'sheet') DEFAULT 'workbook',
                comment TEXT,
                FOREIGN KEY (workbook_id) REFERENCES workbooks(id) ON DELETE CASCADE,
                FOREIGN KEY (sheet_id) REFERENCES sheets(id) ON DELETE CASCADE,
                INDEX idx_name (name)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        ");
        
        // Data validation rules
        $this->pdo->exec("
            CREATE TABLE IF NOT EXISTS data_validations (
                id INT AUTO_INCREMENT PRIMARY KEY,
                sheet_id INT NOT NULL,
                cell_range VARCHAR(255) NOT NULL,
                validation_type VARCHAR(50),
                validation_formula TEXT,
                validation_list JSON,
                error_title VARCHAR(255),
                error_message TEXT,
                show_dropdown BOOLEAN DEFAULT TRUE,
                metadata JSON,
                FOREIGN KEY (sheet_id) REFERENCES sheets(id) ON DELETE CASCADE,
                INDEX idx_sheet_range (sheet_id, cell_range)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        ");
        
        // Cell relationships (for tracking VLOOKUP dependencies)
        $this->pdo->exec("
            CREATE TABLE IF NOT EXISTS cell_relationships (
                id INT AUTO_INCREMENT PRIMARY KEY,
                source_cell_id INT NOT NULL,
                source_type ENUM('independent', 'dependent') NOT NULL,
                target_cell_id INT NOT NULL,
                target_type ENUM('independent', 'dependent') NOT NULL,
                relationship_type VARCHAR(50) COMMENT 'VLOOKUP, MATCH, REFERENCE, etc',
                metadata JSON,
                INDEX idx_source (source_cell_id, source_type),
                INDEX idx_target (target_cell_id, target_type)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        ");
        
        // Audit log for tracking changes
        $this->pdo->exec("
            CREATE TABLE IF NOT EXISTS audit_log (
                id INT AUTO_INCREMENT PRIMARY KEY,
                table_name VARCHAR(50),
                record_id INT,
                action ENUM('INSERT', 'UPDATE', 'DELETE'),
                old_value JSON,
                new_value JSON,
                user VARCHAR(255),
                timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                INDEX idx_table_record (table_name, record_id),
                INDEX idx_timestamp (timestamp)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        ");
    }
    
    /**
     * Import Excel file to MySQL
     */
    public function importExcel($filePath) {
        try {
            // Load the spreadsheet
            $this->spreadsheet = IOFactory::load($filePath);
            
            // Create workbook record
            $stmt = $this->pdo->prepare("
                INSERT INTO workbooks (filename, original_path, metadata) 
                VALUES (?, ?, ?)
            ");
            
            $metadata = json_encode([
                'properties' => [
                    'creator' => $this->spreadsheet->getProperties()->getCreator(),
                    'title' => $this->spreadsheet->getProperties()->getTitle(),
                    'subject' => $this->spreadsheet->getProperties()->getSubject(),
                    'description' => $this->spreadsheet->getProperties()->getDescription(),
                    'keywords' => $this->spreadsheet->getProperties()->getKeywords(),
                    'category' => $this->spreadsheet->getProperties()->getCategory(),
                    'company' => $this->spreadsheet->getProperties()->getCompany(),
                    'created' => $this->spreadsheet->getProperties()->getCreated(),
                    'modified' => $this->spreadsheet->getProperties()->getModified()
                ]
            ]);
            
            $stmt->execute([basename($filePath), $filePath, $metadata]);
            $this->workbookId = $this->pdo->lastInsertId();
            
            // Process each sheet
            foreach ($this->spreadsheet->getSheetNames() as $sheetIndex => $sheetName) {
                $this->processSheet($sheetIndex, $sheetName);
            }
            
            // Process named ranges
            $this->processNamedRanges();
            
            // Build cell relationships
            $this->buildCellRelationships();
            
            echo "Import completed successfully!\n";
            return $this->workbookId;
            
        } catch (Exception $e) {
            $this->pdo->rollBack();
            throw new Exception("Import failed: " . $e->getMessage());
        }
    }
    
    /**
     * Process individual sheet
     */
    private function processSheet($sheetIndex, $sheetName) {
        $sheet = $this->spreadsheet->getSheet($sheetIndex);
        
        // Get sheet dimensions
        $highestRow = $sheet->getHighestRow();
        $highestColumn = $sheet->getHighestColumn();
        $highestColumnIndex = Coordinate::columnIndexFromString($highestColumn);
        
        // Insert sheet record
        $stmt = $this->pdo->prepare("
            INSERT INTO sheets (workbook_id, sheet_name, sheet_index, row_count, column_count, metadata) 
            VALUES (?, ?, ?, ?, ?, ?)
        ");
        
        $sheetMetadata = json_encode([
            'tab_color' => $sheet->getTabColor() ? $sheet->getTabColor()->getARGB() : null,
            'visibility' => $sheet->getSheetState(),
            'right_to_left' => $sheet->getRightToLeft(),
            'protection' => $sheet->getProtection()->isProtectionEnabled()
        ]);
        
        $stmt->execute([
            $this->workbookId, 
            $sheetName, 
            $sheetIndex, 
            $highestRow, 
            $highestColumnIndex,
            $sheetMetadata
        ]);
        
        $sheetId = $this->pdo->lastInsertId();
        
        // Process data validations
        $this->processDataValidations($sheet, $sheetId);
        
        // Process all cells
        foreach ($sheet->getRowIterator() as $row) {
            $rowIndex = $row->getRowIndex();
            
            foreach ($row->getCellIterator() as $cell) {
                if ($cell->getValue() !== null) {
                    $this->processCell($cell, $sheetId, $rowIndex);
                }
            }
        }
    }
    
    /**
     * Process individual cell
     */
    private function processCell($cell, $sheetId, $rowNum) {
        $cellAddress = $cell->getCoordinate();
        $colLetter = preg_replace('/[0-9]/', '', $cellAddress);
        $colNum = Coordinate::columnIndexFromString($colLetter);
        
        // Get cell style information
        $style = $this->extractCellStyle($cell);
        
        // Check if cell has formula
        if ($cell->hasFormula()) {
            $this->processDependentCell($cell, $sheetId, $rowNum, $colNum, $colLetter, $style);
        } else {
            $this->processIndependentCell($cell, $sheetId, $rowNum, $colNum, $colLetter, $style);
        }
    }
    
    /**
     * Process independent (non-formula) cell
     */
    private function processIndependentCell($cell, $sheetId, $rowNum, $colNum, $colLetter, $style) {
        $stmt = $this->pdo->prepare("
            INSERT INTO cells_independent 
            (sheet_id, cell_address, row_num, col_num, col_letter, value, value_type, 
             formatted_value, number_format, hyperlink, comment, style, metadata)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON DUPLICATE KEY UPDATE
                value = VALUES(value),
                value_type = VALUES(value_type),
                formatted_value = VALUES(formatted_value),
                updated_at = CURRENT_TIMESTAMP
        ");
        
        $value = $cell->getValue();
        $formattedValue = $cell->getFormattedValue();
        $valueType = $this->determineValueType($value);
        $numberFormat = $cell->getStyle()->getNumberFormat()->getFormatCode();
        $hyperlink = $cell->hasHyperlink() ? $cell->getHyperlink()->getUrl() : null;
        $comment = $cell->getComment() ? $cell->getComment()->getText()->getPlainText() : null;
        
        $metadata = json_encode([
            'coordinate' => $cell->getCoordinate(),
            'data_type' => $cell->getDataType(),
            'has_rich_text' => ($cell->getValue() instanceof \PhpOffice\PhpSpreadsheet\RichText\RichText)
        ]);
        
        $stmt->execute([
            $sheetId, $cell->getCoordinate(), $rowNum, $colNum, $colLetter,
            $value, $valueType, $formattedValue, $numberFormat, 
            $hyperlink, $comment, json_encode($style), $metadata
        ]);
    }
    
    /**
     * Process dependent (formula) cell
     */
    private function processDependentCell($cell, $sheetId, $rowNum, $colNum, $colLetter, $style) {
        $stmt = $this->pdo->prepare("
            INSERT INTO cells_dependent 
            (sheet_id, cell_address, row_num, col_num, col_letter, formula, formula_type,
             calculated_value, value_type, formatted_value, number_format, dependencies,
             external_references, is_array_formula, has_error, error_message, 
             hyperlink, comment, style, metadata)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON DUPLICATE KEY UPDATE
                formula = VALUES(formula),
                calculated_value = VALUES(calculated_value),
                updated_at = CURRENT_TIMESTAMP
        ");
        
        $formula = $cell->getValue();
        $calculatedValue = $cell->getCalculatedValue();
        $formattedValue = $cell->getFormattedValue();
        $valueType = $this->determineValueType($calculatedValue);
        $numberFormat = $cell->getStyle()->getNumberFormat()->getFormatCode();
        $hyperlink = $cell->hasHyperlink() ? $cell->getHyperlink()->getUrl() : null;
        $comment = $cell->getComment() ? $cell->getComment()->getText()->getPlainText() : null;
        
        // Extract formula type (VLOOKUP, SUM, etc.)
        $formulaType = $this->extractFormulaType($formula);
        
        // Extract dependencies
        $dependencies = $this->extractDependencies($formula);
        $externalRefs = $this->extractExternalReferences($formula);
        
        // Check for errors
        $hasError = is_string($calculatedValue) && strpos($calculatedValue, '#') === 0;
        $errorMessage = $hasError ? $calculatedValue : null;
        
        $metadata = json_encode([
            'coordinate' => $cell->getCoordinate(),
            'data_type' => $cell->getDataType(),
            'is_merged' => $cell->isMergeRangeValueCell()
        ]);
        
        $stmt->execute([
            $sheetId, $cell->getCoordinate(), $rowNum, $colNum, $colLetter,
            $formula, $formulaType, $calculatedValue, $valueType, $formattedValue,
            $numberFormat, json_encode($dependencies), json_encode($externalRefs),
            false, $hasError, $errorMessage, $hyperlink, $comment,
            json_encode($style), $metadata
        ]);
    }
    
    /**
     * Extract cell style information
     */
    private function extractCellStyle($cell) {
        $style = $cell->getStyle();
        
        return [
            'font' => [
                'name' => $style->getFont()->getName(),
                'size' => $style->getFont()->getSize(),
                'bold' => $style->getFont()->getBold(),
                'italic' => $style->getFont()->getItalic(),
                'underline' => $style->getFont()->getUnderline(),
                'color' => $style->getFont()->getColor()->getARGB()
            ],
            'fill' => [
                'type' => $style->getFill()->getFillType(),
                'color' => $style->getFill()->getStartColor()->getARGB()
            ],
            'borders' => [
                'top' => $style->getBorders()->getTop()->getBorderStyle(),
                'bottom' => $style->getBorders()->getBottom()->getBorderStyle(),
                'left' => $style->getBorders()->getLeft()->getBorderStyle(),
                'right' => $style->getBorders()->getRight()->getBorderStyle()
            ],
            'alignment' => [
                'horizontal' => $style->getAlignment()->getHorizontal(),
                'vertical' => $style->getAlignment()->getVertical(),
                'wrap_text' => $style->getAlignment()->getWrapText()
            ]
        ];
    }
    
    /**
     * Determine value type
     */
    private function determineValueType($value) {
        if (is_null($value)) return 'null';
        if (is_bool($value)) return 'boolean';
        if (is_numeric($value)) return 'number';
        if ($value instanceof \DateTime) return 'date';
        return 'string';
    }
    
    /**
     * Extract formula type from formula string
     */
    private function extractFormulaType($formula) {
        preg_match('/^=([A-Z]+)\(/', $formula, $matches);
        return isset($matches[1]) ? $matches[1] : 'CUSTOM';
    }
    
    /**
     * Extract cell dependencies from formula
     */
    private function extractDependencies($formula) {
        $dependencies = [];
        
        // Match cell references (e.g., A1, $B$2, Sheet1!A1)
        preg_match_all('/(?:\'[^\']+\'!)?[$]?[A-Z]+[$]?\d+(?::[$]?[A-Z]+[$]?\d+)?/i', 
                       $formula, $matches);
        
        if (!empty($matches[0])) {
            $dependencies = array_unique($matches[0]);
        }
        
        return $dependencies;
    }
    
    /**
     * Extract external sheet references
     */
    private function extractExternalReferences($formula) {
        $external = [];
        
        // Match sheet references
        preg_match_all("/\'([^']+)\'!/", $formula, $matches);
        
        if (!empty($matches[1])) {
            $external = array_unique($matches[1]);
        }
        
        return $external;
    }
    
    /**
     * Process data validations
     */
    private function processDataValidations($sheet, $sheetId) {
        foreach ($sheet->getDataValidationCollection() as $validation) {
            $stmt = $this->pdo->prepare("
                INSERT INTO data_validations 
                (sheet_id, cell_range, validation_type, validation_formula, 
                 validation_list, error_title, error_message, show_dropdown, metadata)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ");
            
            $validationData = [
                'type' => $validation->getType(),
                'operator' => $validation->getOperator(),
                'allow_blank' => $validation->getAllowBlank(),
                'show_input_message' => $validation->getShowInputMessage(),
                'show_error_message' => $validation->getShowErrorMessage()
            ];
            
            $formula1 = $validation->getFormula1();
            $listValues = null;
            
            // Check if it's a list validation
            if ($validation->getType() == 'list') {
                // Could be a range reference or comma-separated values
                if (strpos($formula1, ',') !== false) {
                    $listValues = explode(',', $formula1);
                }
            }
            
            $stmt->execute([
                $sheetId,
                $validation->getSqref(),
                $validation->getType(),
                $formula1,
                json_encode($listValues),
                $validation->getErrorTitle(),
                $validation->getError(),
                $validation->getShowDropDown(),
                json_encode($validationData)
            ]);
        }
    }
    
    /**
     * Process named ranges
     */
    private function processNamedRanges() {
        $namedRanges = $this->spreadsheet->getNamedRanges();
        
        foreach ($namedRanges as $namedRange) {
            $stmt = $this->pdo->prepare("
                INSERT INTO named_ranges 
                (workbook_id, sheet_id, name, range_address, scope, comment)
                VALUES (?, ?, ?, ?, ?, ?)
            ");
            
            // Get sheet ID if scope is sheet-level
            $sheetId = null;
            if ($namedRange->getScope() !== null) {
                $scopeSheet = $this->spreadsheet->getSheetByName($namedRange->getScope());
                if ($scopeSheet) {
                    // Get sheet ID from database
                    $sheetStmt = $this->pdo->prepare(
                        "SELECT id FROM sheets WHERE workbook_id = ? AND sheet_name = ?"
                    );
                    $sheetStmt->execute([$this->workbookId, $namedRange->getScope()]);
                    $sheetId = $sheetStmt->fetchColumn();
                }
            }
            
            $stmt->execute([
                $this->workbookId,
                $sheetId,
                $namedRange->getName(),
                $namedRange->getValue(),
                $namedRange->getScope() ? 'sheet' : 'workbook',
                $namedRange->getComment()
            ]);
        }
    }
    
    /**
     * Build cell relationships for tracking dependencies
     */
    private function buildCellRelationships() {
        // Get all dependent cells
        $stmt = $this->pdo->prepare("
            SELECT cd.id, cd.sheet_id, cd.dependencies, cd.formula_type
            FROM cells_dependent cd
            WHERE cd.dependencies IS NOT NULL AND cd.dependencies != '[]'
        ");
        $stmt->execute();
        
        while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
            $dependencies = json_decode($row['dependencies'], true);
            
            foreach ($dependencies as $dep) {
                // Parse the dependency reference
                $this->createCellRelationship(
                    $row['id'], 
                    'dependent',
                    $dep, 
                    $row['sheet_id'],
                    $row['formula_type']
                );
            }
        }
    }
    
    /**
     * Create individual cell relationship
     */
    private function createCellRelationship($sourceId, $sourceType, $targetRef, $sheetId, $relType) {
        // Parse target reference to get actual cell
        // This is simplified - you'd need more complex parsing for sheet references
        $targetAddress = preg_replace('/^.*!/', '', $targetRef);
        $targetAddress = str_replace('$', '', $targetAddress);
        
        // Try to find target cell
        $stmt = $this->pdo->prepare("
            SELECT id, 'independent' as type FROM cells_independent 
            WHERE sheet_id = ? AND cell_address = ?
            UNION
            SELECT id, 'dependent' as type FROM cells_dependent 
            WHERE sheet_id = ? AND cell_address = ?
            LIMIT 1
        ");
        $stmt->execute([$sheetId, $targetAddress, $sheetId, $targetAddress]);
        $target = $stmt->fetch(PDO::FETCH_ASSOC);
        
        if ($target) {
            $relStmt = $this->pdo->prepare("
                INSERT IGNORE INTO cell_relationships 
                (source_cell_id, source_type, target_cell_id, target_type, relationship_type)
                VALUES (?, ?, ?, ?, ?)
            ");
            $relStmt->execute([
                $sourceId, 
                $sourceType, 
                $target['id'], 
                $target['type'],
                $relType
            ]);
        }
    }
    
    /**
     * Export data back to Excel (reconstruction)
     */
    public function exportToExcel($workbookId, $outputPath) {
        // This would reconstruct the Excel file from the database
        // Implementation would be quite extensive
        echo "Export functionality would reconstruct Excel from database\n";
    }
    
    /**
     * Update cell value and recalculate dependencies
     */
    public function updateCellValue($sheetId, $cellAddress, $newValue) {
        // Check if cell exists and type
        $stmt = $this->pdo->prepare("
            SELECT 'independent' as type, id FROM cells_independent 
            WHERE sheet_id = ? AND cell_address = ?
            UNION
            SELECT 'dependent' as type, id FROM cells_dependent 
            WHERE sheet_id = ? AND cell_address = ?
        ");
        $stmt->execute([$sheetId, $cellAddress, $sheetId, $cellAddress]);
        $cell = $stmt->fetch(PDO::FETCH_ASSOC);
        
        if ($cell) {
            // Log the change
            $this->logAudit($cell['type'] == 'independent' ? 'cells_independent' : 'cells_dependent',
                           $cell['id'], 'UPDATE', $cellAddress, $newValue);
            
            // Update the value
            if ($cell['type'] == 'independent') {
                $updateStmt = $this->pdo->prepare("
                    UPDATE cells_independent 
                    SET value = ?, value_type = ?, updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                ");
                $updateStmt->execute([$newValue, $this->determineValueType($newValue), $cell['id']]);
            }
            
            // Trigger recalculation of dependent cells
            $this->recalculateDependents($cell['id'], $cell['type']);
        }
    }
    
    /**
     * Recalculate dependent cells
     */
    private function recalculateDependents($cellId, $cellType) {
        // Find all cells that depend on this cell
        $stmt = $this->pdo->prepare("
            SELECT source_cell_id, source_type 
            FROM cell_relationships 
            WHERE target_cell_id = ? AND target_type = ?
        ");
        $stmt->execute([$cellId, $cellType]);
        
        while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
            // Here you would implement formula recalculation
            // This is complex and would require a formula parser/evaluator
            echo "Would recalculate cell ID: {$row['source_cell_id']}\n";
        }
    }
    
    /**
     * Log audit trail
     */
    private function logAudit($table, $recordId, $action, $oldValue, $newValue) {
        $stmt = $this->pdo->prepare("
            INSERT INTO audit_log (table_name, record_id, action, old_value, new_value, user)
            VALUES (?, ?, ?, ?, ?, ?)
        ");
        $stmt->execute([
            $table,
            $recordId,
            $action,
            json_encode($oldValue),
            json_encode($newValue),
            $_SESSION['user'] ?? 'system'
        ]);
    }
    
    /**
     * Get cell information for display
     */
    public function getCellInfo($sheetId, $cellAddress) {
        $sql = "
            SELECT 
                'independent' as cell_type,
                ci.value,
                ci.value_type,
                ci.formatted_value,
                ci.number_format,
                NULL as formula,
                ci.hyperlink,
                ci.comment,
                ci.style,
                s.sheet_name,
                w.filename
            FROM cells_independent ci
            JOIN sheets s ON ci.sheet_id = s.id
            JOIN workbooks w ON s.workbook_id = w.id
            WHERE ci.sheet_id = ? AND ci.cell_address = ?
            
            UNION ALL
            
            SELECT 
                'dependent' as cell_type,
                cd.calculated_value as value,
                cd.value_type,
                cd.formatted_value,
                cd.number_format,
                cd.formula,
                cd.hyperlink,
                cd.comment,
                cd.style,
                s.sheet_name,
                w.filename
            FROM cells_dependent cd
            JOIN sheets s ON cd.sheet_id = s.id
            JOIN workbooks w ON s.workbook_id = w.id
            WHERE cd.sheet_id = ? AND cd.cell_address = ?
        ";
        
        $stmt = $this->pdo->prepare($sql);
        $stmt->execute([$sheetId, $cellAddress, $sheetId, $cellAddress]);
        return $stmt->fetch(PDO::FETCH_ASSOC);
    }
    
    /**
     * Create view for easy phpMyAdmin browsing
     */
    public function createUserFriendlyViews() {
        // Combined cell view
        $this->pdo->exec("
            CREATE OR REPLACE VIEW v_all_cells AS
            SELECT 
                w.filename as workbook,
                s.sheet_name as sheet,
                ci.cell_address,
                ci.row_num,
                ci.col_letter,
                'Value' as cell_type,
                ci.value,
                NULL as formula,
                ci.formatted_value,
                ci.comment
            FROM cells_independent ci
            JOIN sheets s ON ci.sheet_id = s.id
            JOIN workbooks w ON s.workbook_id = w.id
            
            UNION ALL
            
            SELECT 
                w.filename as workbook,
                s.sheet_name as sheet,
                cd.cell_address,
                cd.row_num,
                cd.col_letter,
                'Formula' as cell_type,
                cd.calculated_value as value,
                cd.formula,
                cd.formatted_value,
                cd.comment
            FROM cells_dependent cd
            JOIN sheets s ON cd.sheet_id = s.id
            JOIN workbooks w ON s.workbook_id = w.id
            ORDER BY workbook, sheet, row_num, col_letter
        ");
        
        // VLOOKUP analysis view
        $this->pdo->exec("
            CREATE OR REPLACE VIEW v_vlookup_analysis AS
            SELECT 
                w.filename as workbook,
                s.sheet_name as sheet,
                cd.cell_address,
                cd.formula,
                cd.calculated_value,
                cd.dependencies,
                cd.external_references,
                cd.has_error,
                cd.error_message
            FROM cells_dependent cd
            JOIN sheets s ON cd.sheet_id = s.id
            JOIN workbooks w ON s.workbook_id = w.id
            WHERE cd.formula_type = 'VLOOKUP'
        ");
        
        // Data validation dropdowns view
        $this->pdo->exec("
            CREATE OR REPLACE VIEW v_dropdown_cells AS
            SELECT 
                w.filename as workbook,
                s.sheet_name as sheet,
                dv.cell_range,
                dv.validation_type,
                dv.validation_formula,
                dv.validation_list,
                dv.error_title,
                dv.error_message
            FROM data_validations dv
            JOIN sheets s ON dv.sheet_id = s.id
            JOIN workbooks w ON s.workbook_id = w.id
            WHERE dv.validation_type = 'list'
        ");
    }
}

// Usage example
try {
    $converter = new ExcelToMySQLConverter('localhost', 'excel_db', 'username', 'password');
    
    // Import Excel file
    $workbookId = $converter->importExcel('path/to/your/excel_file.xlsx');
    
    // Create user-friendly views for phpMyAdmin
    $converter->createUserFriendlyViews();
    
    // Example: Get cell information
    $cellInfo = $converter->getCellInfo(1, 'A1');
    print_r($cellInfo);
    
    // Example: Update a cell value
    $converter->updateCellValue(1, 'B2', 'New Value');
    
} catch (Exception $e) {
    echo "Error: " . $e->getMessage();
}