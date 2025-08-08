<?php
header('Content-Type: application/json');

try {
    $input = json_decode(file_get_contents('php://input'), true);
    
    require_once __DIR__ . '/../config/database.php';
    $config = require __DIR__ . '/../config/database.php';
    
    $pdo = new PDO(
        "mysql:host={$config['host']};dbname={$config['database']};charset={$config['charset']}",
        $config['username'],
        $config['password']
    );
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    
    $sheetId = $input['sheet_id'];
    
    // Get all formula cells
    $stmt = $pdo->prepare("
        SELECT cell_address 
        FROM cells_dependent 
        WHERE sheet_id = ?
    ");
    $stmt->execute([$sheetId]);
    $allFormulaCells = $stmt->fetchAll(PDO::FETCH_COLUMN);
    
    // Recalculate all
    $_POST = json_encode(['sheet_id' => $sheetId, 'cells' => $allFormulaCells]);
    include 'recalculate.php';
    
} catch (Exception $e) {
    echo json_encode(['error' => $e->getMessage()]);
}