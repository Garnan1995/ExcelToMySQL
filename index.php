<?php
session_start();
error_reporting(E_ALL);
ini_set('display_errors', 1);

// ADD THESE LINES:
set_time_limit(0);  // No time limit
ini_set('memory_limit', '512M');  // Increase memory

// Autoload dependencies
require_once 'vendor/autoload.php';
require_once 'src/ExcelToMySQLConverter.php';
require_once 'src/FormulaEvaluator.php';
require_once 'src/ExcelDatabaseUI.php';

// Load database configuration
$config = require 'config/database.php';

// Create database connection
try {
    $pdo = new PDO(
        "mysql:host={$config['host']};port={$config['port']};charset={$config['charset']}",
        $config['username'],
        $config['password']
    );
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    
    // Create database if not exists
    $pdo->exec("CREATE DATABASE IF NOT EXISTS {$config['database']}");
    $pdo->exec("USE {$config['database']}");
    
    echo "<h1>Excel to MySQL Converter</h1>";
    echo "<p style='color: green;'>✓ Database connection successful!</p>";
    
} catch (PDOException $e) {
    die("<p style='color: red;'>✗ Database connection failed: " . $e->getMessage() . "</p>");
}

// Initialize converter
$converter = new ExcelToMySQLConverter(
    $config['host'],
    $config['database'],
    $config['username'],
    $config['password']
);

// Handle file upload
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['excel_file'])) {
    if ($_FILES['excel_file']['error'] === UPLOAD_ERR_OK) {
        $uploadDir = __DIR__ . '/uploads/';
        $uploadFile = $uploadDir . basename($_FILES['excel_file']['name']);
        
        if (move_uploaded_file($_FILES['excel_file']['tmp_name'], $uploadFile)) {
            try {
                $workbookId = $converter->importExcel($uploadFile);
                echo "<p style='color: green;'>✓ File imported successfully! Workbook ID: $workbookId</p>";
                
                // Create user-friendly views
                $converter->createUserFriendlyViews();
                echo "<p style='color: green;'>✓ Database views created!</p>";
                
            } catch (Exception $e) {
                echo "<p style='color: red;'>✗ Import failed: " . $e->getMessage() . "</p>";
            }
        } else {
            echo "<p style='color: red;'>✗ File upload failed!</p>";
        }
    }
}
?>

<!DOCTYPE html>
<html>
<head>
    <title>Excel to MySQL Converter</title>
    <style>
        body { 
            font-family: Arial, sans-serif; 
            max-width: 1200px; 
            margin: 0 auto; 
            padding: 20px;
        }
        .upload-form {
            background: #f0f0f0;
            padding: 20px;
            border-radius: 5px;
            margin: 20px 0;
        }
        .info-box {
            background: #e8f4ff;
            padding: 15px;
            border-left: 4px solid #0066cc;
            margin: 20px 0;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin: 20px 0;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        th {
            background: #f0f0f0;
        }
    </style>
</head>
<body>
    <div class="upload-form">
        <h2>Upload Excel File</h2>
        <form method="POST" enctype="multipart/form-data">
            <input type="file" name="excel_file" accept=".xlsx,.xls" required>
            <button type="submit">Import to Database</button>
        </form>
    </div>
    
    <div class="info-box">
        <h3>Quick Links:</h3>
        <ul>
            <li><a href="http://localhost/phpmyadmin" target="_blank">Open phpMyAdmin</a></li>
            <li><a href="view_sheets.php">View Imported Sheets</a></li>
            <li><a href="view_formulas.php">View All Formulas</a></li>
        </ul>
    </div>
    
    <?php
    // Display existing workbooks
    try {
        $stmt = $pdo->query("SELECT * FROM workbooks ORDER BY import_date DESC");
        if ($stmt->rowCount() > 0) {
            echo "<h2>Imported Workbooks</h2>";
            echo "<table>";
            echo "<tr><th>ID</th><th>Filename</th><th>Import Date</th><th>Actions</th></tr>";
            
            while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
                echo "<tr>";
                echo "<td>{$row['id']}</td>";
                echo "<td>{$row['filename']}</td>";
                echo "<td>{$row['import_date']}</td>";
                echo "<td><a href='view_workbook.php?id={$row['id']}'>View Sheets</a></td>";
                echo "</tr>";
            }
            echo "</table>";
        }
    } catch (PDOException $e) {
        // Tables might not exist yet
    }
    ?>
</body>
</html>