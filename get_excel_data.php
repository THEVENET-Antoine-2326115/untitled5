<?php
// get_excel_data.php
ini_set('display_errors', 0);
error_reporting(0);

// Inclure le nouveau processeur Excel
require_once 'ExcelProcessor.php';

// Permettre les requêtes AJAX cross-origin si nécessaire
header('Content-Type: application/json');

// Chemins vers les fichiers Excel
$excelFilePaths = [
    'file1' => "BPU_vrai.xlsx",
    'file2' => "ref.xlsx",
    'file3' => "NomenclatureEntretoise.xlsx"
];

$response = [];

// Traiter chaque fichier Excel
foreach ($excelFilePaths as $fileKey => $filePath) {
    // Création de l'objet ExcelProcessor
    $excelProcessor = new ExcelProcessor($filePath);

    // Traitement des données Excel
    $success = $excelProcessor->processData();

    if ($success) {
        // Récupération des données et des totaux par produit
        $response[$fileKey] = [
            'data' => $excelProcessor->getData(),
            'productTotals' => $excelProcessor->getProductTotals(),
            'filePath' => $excelProcessor->getFilePath()
        ];
    } else {
        // En cas d'erreur pour ce fichier
        $response[$fileKey] = ['error' => "Erreur lors de la lecture du fichier Excel: $filePath"];
    }
}

// Conversion en JSON
echo json_encode($response);