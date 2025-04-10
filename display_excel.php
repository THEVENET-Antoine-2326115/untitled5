<?php
// display_excel.php
require_once 'ExcelReader.php';

// Chemin vers le fichier Excel
$excelFilePath = "test.xlsx";

// Création de l'objet ExcelReader
$excelReader = new ExcelReader($excelFilePath);

// Lecture du fichier Excel
$success = $excelReader->read();

// Récupération des données
$data = $excelReader->getData();
?>