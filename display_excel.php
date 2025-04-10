<?php
// display_excel.php
require_once 'ExcelReader.php';

// Chemin vers le fichier Excel
$excelFilePath = "BPU_Moscatelli entretoises harmonisées 25-03-28.xlsx";

// Création de l'objet ExcelReader
$excelReader = new ExcelReader($excelFilePath);

// Lecture du fichier Excel
$success = $excelReader->read();

// Récupération des données
$data = $excelReader->getData();
?>