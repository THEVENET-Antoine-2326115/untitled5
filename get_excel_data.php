<?php

// get_excel_data.php
ini_set('display_errors', 0);
error_reporting(0);


// get_excel_data.php
require_once 'ExcelReader.php';

// Permettre les requêtes AJAX cross-origin si nécessaire
header('Content-Type: application/json');

// Chemin vers le fichier Excel
$excelFilePath = "test.xlsx";

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