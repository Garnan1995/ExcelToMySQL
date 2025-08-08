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
    
    $workbookId = $_GET['workbook_id'] ?? 0;
    
    $stmt = $pdo->prepare("
        SELECT id, sheet_name, sheet_index 
        FROM sheets 
        WHERE workbook_id = ? 
        ORDER BY sheet_index
    ");
    $stmt->execute([$workbookId]);
    echo json_encode($stmt->fetchAll(PDO::FETCH_ASSOC));
    
} catch (Exception $e) {
    echo json_encode(['error' => $e->getMessage()]);
}