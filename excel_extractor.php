<?php
/**
 * Excel to JSON and MySQL Extraction - Optimized Version
 * Requires: composer require phpoffice/phpspreadsheet
 */

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataValidation;
use PhpOffice\PhpSpreadsheet\Shared\Date;

// Increase memory limit and execution time for large files
ini_set('memory_limit', '512M');
ini_set('max_execution_time', 300);

class ExcelExtractor {
    private $excelFile;
    private $sheetName;
    private $dbHost = 'localhost';
    private $dbUser = 'root';
    private $dbPass = '';
    private $dbName = 'my_excel_data_2';
    private $tableName;
    private $conn;
    private $debug = true; // Enable debug output
    
    public function __construct($excelFile = 'HPP.xlsx', $sheetName = 'HPP') {
        $this->excelFile = $excelFile;
        $this->sheetName = $sheetName;
        $this->tableName = $this->sanitizeName($sheetName);
    }
    
    /**
     * Sanitize names for MySQL tables/columns
     */
    private function sanitizeName($name, $maxLength = 64) {
        $name = preg_replace('/[^a-zA-Z0-9_]/', '_', $name);
        $name = preg_replace('/_+/', '_', $name);
        $name = trim($name, '_');
        $name = strtolower($name);
        
        if (strlen($name) > $maxLength) {
            $name = substr($name, 0, 30) . '_' . substr($name, -($maxLength - 31));
        }
        
        return $name;
    }
    
    /**
     * Extract all cell information from Excel - Simplified version
     */
    public function extractCellData() {
        echo "========================================\n";
        echo "Excel Cell Data Extractor (Optimized)\n";
        echo "========================================\n\n";
        
        try {
            // Check if file exists
            if (!file_exists($this->excelFile)) {
                throw new Exception("Excel file not found: {$this->excelFile}");
            }
            
            echo "Loading Excel file: {$this->excelFile}\n";
            
            // Load with read data only to improve performance
            $reader = IOFactory::createReader('Xlsx');
            $reader->setReadDataOnly(false); // We need formatting
            $reader->setReadEmptyCells(false); // Skip empty cells
            
            $spreadsheet = $reader->load($this->excelFile);
            
            // Get the specific sheet
            $worksheet = $spreadsheet->getSheetByName($this->sheetName);
            if (!$worksheet) {
                throw new Exception("Sheet '{$this->sheetName}' not found in Excel file");
            }
            
            echo "Processing sheet: {$this->sheetName}\n";
            echo "----------------------------------------\n";
            
            // Get actual used range
            $highestRow = $worksheet->getHighestDataRow();
            $highestColumn = $worksheet->getHighestDataColumn();
            $highestColumnIndex = Coordinate::columnIndexFromString($highestColumn);
            
            echo "Sheet dimensions: {$highestColumn}{$highestRow} ";
            echo "({$highestColumnIndex} columns × {$highestRow} rows)\n";
            echo "Starting cell extraction...\n\n";
            
            $cellData = [];
            $processedCells = 0;
            $totalCells = $highestRow * $highestColumnIndex;
            
            // Process cells row by row for better memory management
            for ($row = 1; $row <= $highestRow; $row++) {
                if ($this->debug && $row % 10 == 0) {
                    $percent = round(($row / $highestRow) * 100);
                    echo "Processing row {$row}/{$highestRow} ({$percent}%)...\r";
                }
                
                for ($col = 1; $col <= $highestColumnIndex; $col++) {
                    $columnLetter = Coordinate::stringFromColumnIndex($col);
                    $cellAddress = $columnLetter . $row;
                    
                    try {
                        // Get cell
                        $cell = $worksheet->getCell($cellAddress, false);
                        
                        // Skip truly empty cells
                        $value = $cell->getValue();
                        if ($value === null || $value === '') {
                            continue;
                        }
                        
                        // Build cell info - simplified for performance
                        $cellInfo = [
                            'address' => $cellAddress,
                            'row' => $row,
                            'column' => $columnLetter,
                            'value' => $this->getCellValueSimple($cell),
                            'formula' => null,
                            'dataType' => $cell->getDataType()
                        ];
                        
                        // Check if it's a formula
                        if ($cell->isFormula()) {
                            $cellInfo['formula'] = $value;
                            // Get calculated value
                            try {
                                $cellInfo['value'] = $cell->getCalculatedValue();
                            } catch (Exception $e) {
                                $cellInfo['value'] = '#ERROR';
                            }
                        }
                        
                        // Add basic formatting (simplified)
                        $cellInfo['formatting'] = $this->getCellFormattingSimple($worksheet, $cellAddress);
                        
                        // Add validation if exists (simplified check)
                        $cellInfo['validation'] = $this->getCellValidationSimple($worksheet, $cellAddress);
                        
                        $cellData[] = $cellInfo;
                        $processedCells++;
                        
                    } catch (Exception $e) {
                        if ($this->debug) {
                            echo "\nWarning: Error processing cell {$cellAddress}: " . $e->getMessage() . "\n";
                        }
                        continue;
                    }
                }
            }
            
            echo "\n\nTotal cells with data: {$processedCells}\n";
            echo "----------------------------------------\n\n";
            
            // Save to JSON
            $this->saveToJson($cellData);
            
            // Save to MySQL
            $this->saveToMySQL($cellData);
            
            // Create summary
            $this->createSummary($cellData);
            
            // Free memory
            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);
            
            return $cellData;
            
        } catch (Exception $e) {
            echo "\nError: " . $e->getMessage() . "\n";
            echo "Stack trace:\n" . $e->getTraceAsString() . "\n";
            return false;
        }
    }
    
    /**
     * Get cell value - simplified version
     */
    private function getCellValueSimple($cell) {
        try {
            $value = $cell->getCalculatedValue();
            
            // Handle dates
            if (Date::isDateTime($cell) && is_numeric($value)) {
                $dateObj = Date::excelToDateTimeObject($value);
                return $dateObj->format('Y-m-d H:i:s');
            }
            
            return $value;
        } catch (Exception $e) {
            return $cell->getValue();
        }
    }
    
    /**
     * Get cell formatting - simplified for performance
     */
    private function getCellFormattingSimple($worksheet, $cellAddress) {
        try {
            $style = $worksheet->getStyle($cellAddress);
            $formatting = [];
            
            // Only get number format
            $numberFormat = $style->getNumberFormat()->getFormatCode();
            if ($numberFormat && $numberFormat != 'General') {
                $formatting['numberFormat'] = $numberFormat;
            }
            
            // Basic font info
            $font = $style->getFont();
            if ($font->getBold()) {
                $formatting['bold'] = true;
            }
            
            return empty($formatting) ? null : $formatting;
            
        } catch (Exception $e) {
            return null;
        }
    }
    
    /**
     * Get cell validation - simplified
     */
    private function getCellValidationSimple($worksheet, $cellAddress) {
        try {
            $validation = $worksheet->getCell($cellAddress)->getDataValidation();
            
            if (!$validation || $validation->getType() == DataValidation::TYPE_NONE) {
                return null;
            }
            
            return [
                'type' => $validation->getType(),
                'formula1' => $validation->getFormula1()
            ];
        } catch (Exception $e) {
            return null;
        }
    }
    
    /**
     * Save extracted data to JSON file
     */
    private function saveToJson($cellData) {
        echo "Saving to JSON...\n";
        
        $jsonFile = pathinfo($this->excelFile, PATHINFO_FILENAME) . '_extracted.json';
        
        $jsonData = [
            'metadata' => [
                'sourceFile' => $this->excelFile,
                'sheetName' => $this->sheetName,
                'extractionDate' => date('Y-m-d H:i:s'),
                'totalCells' => count($cellData)
            ],
            'cells' => $cellData
        ];
        
        $json = json_encode($jsonData, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE);
        
        if (file_put_contents($jsonFile, $json) === false) {
            throw new Exception("Failed to write JSON file");
        }
        
        $fileSize = number_format(filesize($jsonFile) / 1024, 2);
        echo "✓ Saved to JSON: {$jsonFile} ({$fileSize} KB)\n\n";
    }
    
    /**
     * Save extracted data to MySQL
     */
    private function saveToMySQL($cellData) {
        echo "Saving to MySQL...\n";
        
        try {
            // Connect to MySQL
            $this->conn = new mysqli($this->dbHost, $this->dbUser, $this->dbPass);
            
            if ($this->conn->connect_error) {
                throw new Exception("Connection failed: " . $this->conn->connect_error);
            }
            
            // Create database if not exists
            $this->conn->query("CREATE DATABASE IF NOT EXISTS {$this->dbName}");
            $this->conn->select_db($this->dbName);
            
            echo "Connected to database: {$this->dbName}\n";
            
            // Drop existing table
            $this->conn->query("DROP TABLE IF EXISTS `{$this->tableName}_cells`");
            
            // Create simplified table structure
            $createTable = "
            CREATE TABLE `{$this->tableName}_cells` (
                id INT AUTO_INCREMENT PRIMARY KEY,
                cell_address VARCHAR(10) NOT NULL,
                row_num INT NOT NULL,
                column_letter VARCHAR(5) NOT NULL,
                cell_value TEXT,
                formula TEXT,
                data_type VARCHAR(20),
                formatting TEXT,
                validation TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                INDEX idx_address (cell_address),
                INDEX idx_row_col (row_num, column_letter)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            ";
            
            if (!$this->conn->query($createTable)) {
                throw new Exception("Error creating table: " . $this->conn->error);
            }
            
            echo "Created table: {$this->tableName}_cells\n";
            
            // Insert data using multi-insert for better performance
            $values = [];
            $inserted = 0;
            $batchSize = 100;
            
            foreach ($cellData as $cell) {
                // Prepare values
                $cellValue = is_array($cell['value']) || is_object($cell['value']) 
                    ? json_encode($cell['value']) 
                    : (string)$cell['value'];
                
                $formatting = $cell['formatting'] ? json_encode($cell['formatting']) : null;
                $validation = $cell['validation'] ? json_encode($cell['validation']) : null;
                
                // Escape values
                $values[] = sprintf(
                    "('%s', %d, '%s', '%s', %s, '%s', %s, %s)",
                    $this->conn->real_escape_string($cell['address']),
                    $cell['row'],
                    $this->conn->real_escape_string($cell['column']),
                    $this->conn->real_escape_string($cellValue),
                    $cell['formula'] ? "'" . $this->conn->real_escape_string($cell['formula']) . "'" : "NULL",
                    $this->conn->real_escape_string($cell['dataType']),
                    $formatting ? "'" . $this->conn->real_escape_string($formatting) . "'" : "NULL",
                    $validation ? "'" . $this->conn->real_escape_string($validation) . "'" : "NULL"
                );
                
                // Insert in batches
                if (count($values) >= $batchSize) {
                    $sql = "INSERT INTO `{$this->tableName}_cells` 
                            (cell_address, row_num, column_letter, cell_value, formula, data_type, formatting, validation) 
                            VALUES " . implode(',', $values);
                    
                    if ($this->conn->query($sql)) {
                        $inserted += count($values);
                        echo "Inserted {$inserted}/" . count($cellData) . " cells\r";
                    } else {
                        echo "\nWarning: Batch insert failed: " . $this->conn->error . "\n";
                    }
                    
                    $values = [];
                }
            }
            
            // Insert remaining values
            if (!empty($values)) {
                $sql = "INSERT INTO `{$this->tableName}_cells` 
                        (cell_address, row_num, column_letter, cell_value, formula, data_type, formatting, validation) 
                        VALUES " . implode(',', $values);
                
                if ($this->conn->query($sql)) {
                    $inserted += count($values);
                }
            }
            
            echo "\n✓ Successfully inserted {$inserted} cells into MySQL\n\n";
            
        } catch (Exception $e) {
            echo "MySQL Error: " . $e->getMessage() . "\n";
        }
    }
    
    /**
     * Create extraction summary
     */
    private function createSummary($cellData) {
        echo "========================================\n";
        echo "EXTRACTION COMPLETE!\n";
        echo "========================================\n";
        
        // Calculate statistics
        $stats = [
            'Total Cells' => count($cellData),
            'Cells With Formulas' => 0,
            'Cells With Validation' => 0,
            'Cells With Formatting' => 0
        ];
        
        foreach ($cellData as $cell) {
            if (!empty($cell['formula'])) $stats['Cells With Formulas']++;
            if (!empty($cell['validation'])) $stats['Cells With Validation']++;
            if (!empty($cell['formatting'])) $stats['Cells With Formatting']++;
        }
        
        echo "\nExtraction Statistics:\n";
        echo "----------------------\n";
        foreach ($stats as $label => $value) {
            echo sprintf("%-25s: %d\n", $label, $value);
        }
        
        echo "\nOutput Files:\n";
        echo "-------------\n";
        echo "✓ JSON: " . pathinfo($this->excelFile, PATHINFO_FILENAME) . "_extracted.json\n";
        echo "✓ MySQL Database: {$this->dbName}\n";
        echo "✓ MySQL Table: {$this->tableName}_cells\n";
        echo "\nView in phpMyAdmin: http://localhost/phpmyadmin → {$this->dbName}\n";
        echo "========================================\n";
    }
}

// Run the extraction with error reporting
error_reporting(E_ALL);

try {
    echo "Starting extraction process...\n\n";
    $extractor = new ExcelExtractor('HPP.xlsx', 'HPP');
    $result = $extractor->extractCellData();
    
    if ($result === false) {
        echo "\nExtraction failed. Please check the error messages above.\n";
    }
} catch (Exception $e) {
    echo "\nFatal error: " . $e->getMessage() . "\n";
    echo "Stack trace:\n" . $e->getTraceAsString() . "\n";
}

?>