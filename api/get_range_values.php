<?php
header('Content-Type: application/json');

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
    
    // Remove $ signs from range
    $range = str_replace('$', '', $range);
    
    // Parse range (e.g., R3:R42)
    if (preg_match('/([A-Z]+)(\d+):([A-Z]+)(\d+)/', $range, $matches)) {
        $startCol = $matches[1];
        $startRow = $matches[2];
        $endCol = $matches[3];
        $endRow = $matches[4];
        
        // Get sheet ID
        $stmt = $pdo->prepare("SELECT id FROM sheets WHERE sheet_name = ?");
        $stmt->execute([$sheetName]);
        $sheetId = $stmt->fetchColumn();
        
        if ($sheetId) {
            // Get values from range
            $stmt = $pdo->prepare("
                SELECT DISTINCT value 
                FROM cells_independent 
                WHERE sheet_id = ? 
                AND col_letter = ? 
                AND row_num BETWEEN ? AND ?
                AND value IS NOT NULL 
                AND value != ''
                ORDER BY row_num
            ");
            $stmt->execute([$sheetId, $startCol, $startRow, $endRow]);
            
            $values = $stmt->fetchAll(PDO::FETCH_COLUMN);
            echo json_encode(['values' => $values]);
        } else {
            echo json_encode(['values' => []]);
        }
    } else {
        echo json_encode(['values' => []]);
    }
    
} catch (Exception $e) {
    echo json_encode(['error' => $e->getMessage(), 'values' => []]);
}