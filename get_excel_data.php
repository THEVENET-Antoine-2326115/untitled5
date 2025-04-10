<?php
// get_excel_data.php
ini_set('display_errors', 0);
error_reporting(0);

// Inclure le nouveau processeur Excel
require_once 'ExcelProcessor.php';

// Permettre les requêtes AJAX cross-origin si nécessaire
header('Content-Type: application/json');

// Chemin vers le fichier Excel
$excelFilePath = "test.xlsx";

// Création de l'objet ExcelProcessor
$excelProcessor = new ExcelProcessor($excelFilePath);

// Traitement des données Excel
$success = $excelProcessor->processData();

if ($success) {
    // Récupération des données et des totaux
    $response = [
        'data' => $excelProcessor->getData(),
        'totals' => $excelProcessor->getTotals()
    ];

    // Conversion en JSON
    echo json_encode($response);
} else {
    // En cas d'erreur
    echo json_encode(['error' => 'Erreur lors de la lecture du fichier Excel']);
}