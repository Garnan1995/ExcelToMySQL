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
        SELECT * FROM data_validations 
        WHERE sheet_id = ?
    ");
    $stmt->execute([$sheetId]);
    echo json_encode($stmt->fetchAll(PDO::FETCH_ASSOC));
    
} catch (Exception $e) {
    echo json_encode(['error' => $e->getMessage()]);
}