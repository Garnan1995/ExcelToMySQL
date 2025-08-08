<?php
require_once 'vendor/autoload.php';
require_once 'src/ExcelDatabaseUI.php';

$config = require 'config/database.php';

try {
    $pdo = new PDO(
        "mysql:host={$config['host']};dbname={$config['database']};charset={$config['charset']}",
        $config['username'],
        $config['password']
    );
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    
    $ui = new ExcelDatabaseUI($pdo);
    
    if (isset($_GET['sheet_id'])) {
        // Display specific sheet
        echo $ui->generateSheetView($_GET['sheet_id']);
    } elseif (isset($_GET['id'])) {
        // Display workbook sheets
        $stmt = $pdo->prepare("
            SELECT * FROM sheets WHERE workbook_id = ? ORDER BY sheet_index
        ");
        $stmt->execute([$_GET['id']]);
        
        echo "<h2>Workbook Sheets</h2>";
        echo "<ul>";
        while ($sheet = $stmt->fetch(PDO::FETCH_ASSOC)) {
            echo "<li><a href='?sheet_id={$sheet['id']}'>{$sheet['sheet_name']}</a></li>";
        }
        echo "</ul>";
    }
    
} catch (PDOException $e) {
    echo "Error: " . $e->getMessage();
}
?>