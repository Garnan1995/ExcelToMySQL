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
    
    $sheetId = $_GET['sheet_id'] ?? 0;
    
    $stmt = $pdo->prepare("
        SELECT 
            cell_address, value, NULL as formula, value_type,
            formatted_value, NULL as dependencies, NULL as error_message
        FROM cells_independent 
        WHERE sheet_id = ?
        
        UNION ALL
        
        SELECT 
            cell_address, calculated_value as value, formula, value_type,
            formatted_value, dependencies, error_message
        FROM cells_dependent 
        WHERE sheet_id = ?
    ");
    $stmt->execute([$sheetId, $sheetId]);
    echo json_encode($stmt->fetchAll(PDO::FETCH_ASSOC));
    
} catch (Exception $e) {
    echo json_encode(['error' => $e->getMessage()]);
}