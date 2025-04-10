<?php
// get_excel_data.php
require_once 'ExcelReader.php';

// Permettre les requêtes AJAX cross-origin si nécessaire
header('Content-Type: application/json');

// Chemin vers le fichier Excel
$excelFilePath = "BPU_Moscatelli entretoises harmonisées 25-03-28.xlsx";

// Création de l'objet ExcelReader
$excelReader = new ExcelReader($excelFilePath);

// Lecture du fichier Excel
$success = $excelReader->read();

if ($success) {
    // Récupération des données et conversion en JSON
    echo json_encode($excelReader->getData());
} else {
    // En cas d'erreur
    echo json_encode(['error' => 'Erreur lors de la lecture du fichier Excel']);
}