<?php
// export_excel.php

// Activer l'affichage des erreurs (à désactiver en production)
ini_set('display_errors', 0);
error_reporting(0);

// Inclure les dépendances
require 'vendor/autoload.php';

// Utiliser les classes Spout pour la manipulation des fichiers Excel
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Writer\Common\Creator\Style\StyleBuilder;
use Box\Spout\Common\Entity\Style\Color;
use Box\Spout\Common\Entity\Style\Border;
use Box\Spout\Common\Entity\Style\BorderBuilder;
use Box\Spout\Common\Entity\Row;
use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;

// Récupérer les données JSON envoyées par le client
$jsonData = file_get_contents('php://input');
$data = json_decode($jsonData, true);

// Vérifier si les données sont valides
if (!$data || !isset($data['recap']) || !isset($data['prices'])) {
    header('HTTP/1.1 400 Bad Request');
    echo json_encode(['error' => 'Données invalides']);
    exit;
}

// Chemins des fichiers
$bpuFilePath = "BPU_vrai.xlsx";
$tempFilePath = "temp_export_" . uniqid() . ".xlsx";

try {
    // Créer un nouveau fichier Excel
    $writer = WriterEntityFactory::createXLSXWriter();
    $writer->openToFile($tempFilePath);

    // Créer des styles pour les en-têtes et les cellules
    $headerStyle = (new StyleBuilder())
        ->setFontBold()
        ->setFontSize(12)
        ->setBackgroundColor(Color::LIGHT_BLUE)
        ->setFontColor(Color::BLACK)
        ->build();

    $totalRowStyle = (new StyleBuilder())
        ->setFontBold()
        ->setFontSize(12)
        ->setBackgroundColor(Color::LIGHT_GRAY)
        ->build();

    $priceStyle = (new StyleBuilder())
        ->setFontBold()
        ->setFontSize(12)
        ->build();

    // ======= FEUILLE 1 : RÉCAPITULATIF =======

    // Ajouter la feuille de récapitulatif
    $writer->addRow(WriterEntityFactory::createRowFromArray(['TABLEAU RÉCAPITULATIF'], $headerStyle));
    $writer->addRow(WriterEntityFactory::createRowFromArray([''])); // Ligne vide

    // En-têtes du tableau récapitulatif
    $recapHeaders = [
        'Site',
        'Repère de panneaux',
        'Quantité',
        'Transport Barreaux',
        'Manutention Barreaux',
        'Nbr Demi-Journées',
        'Grutage Barreaux',
        'Transport Panneaux',
        'Manutention Panneaux',
        'Grutage Panneaux',
        'Emballage',
        'Prix Total'
    ];
    $writer->addRow(WriterEntityFactory::createRowFromArray($recapHeaders, $headerStyle));

    // Données du tableau récapitulatif
    foreach ($data['recap'] as $item) {
        $rowData = [
            $item['site'] ?? '',
            $item['description'] ?? '',
            $item['quantity'] ?? '',
            $item['transportBarreaux'] ?? '',
            $item['manutentionBarreaux'] ?? '',
            $item['manutentionDays'] ?? '',
            $item['dechargementBarreaux'] ?? '',
            $item['transportPanneaux'] ?? '',
            $item['manutentionPanneaux'] ?? '',
            $item['grutagePanneaux'] ?? '',
            $item['emballage'] ?? '',
            isset($item['totalPrice']) ? number_format($item['totalPrice'], 2, ',', ' ') . ' €' : ''
        ];
        $writer->addRow(WriterEntityFactory::createRowFromArray($rowData));
    }

    // Ajouter un espace
    $writer->addRow(WriterEntityFactory::createRowFromArray(['']));
    $writer->addRow(WriterEntityFactory::createRowFromArray(['']));

    // Ajouter le récapitulatif des prix
    $writer->addRow(WriterEntityFactory::createRowFromArray(['RÉCAPITULATIF DES PRIX'], $headerStyle));
    $writer->addRow(WriterEntityFactory::createRowFromArray([''])); // Ligne vide

    $priceData = [
        ['Forfaits (logistique et suivi)', number_format($data['prices']['prixForfaits'], 2, ',', ' ') . ' €'],
        ['Soudage des barreaux', number_format($data['prices']['prixSoudage'], 2, ',', ' ') . ' €'],
        ['Transport des barreaux', number_format($data['prices']['prixTransportBarreaux'], 2, ',', ' ') . ' €'],
        ['Transport des panneaux', number_format($data['prices']['prixTransportPanneaux'], 2, ',', ' ') . ' €'],
        ['Manutention des barreaux', number_format($data['prices']['prixManutention'], 2, ',', ' ') . ' €'],
        ['Manutention des panneaux', number_format($data['prices']['prixManutentionPanneaux'], 2, ',', ' ') . ' €'],
        ['Grutage des barreaux', number_format($data['prices']['prixGrutage'], 2, ',', ' ') . ' €'],
        ['Grutage des panneaux', number_format($data['prices']['prixGrutagePanneaux'], 2, ',', ' ') . ' €'],
        ['Emballage', number_format($data['prices']['prixEmballage'], 2, ',', ' ') . ' €'],
        ['Montage des panneaux', number_format($data['prices']['prixMontage'], 2, ',', ' ') . ' €'],
        ['PRIX TOTAL', number_format($data['prices']['prixTotal'], 2, ',', ' ') . ' €']
    ];

    foreach ($priceData as $index => $row) {
        // Utiliser le style total pour la dernière ligne (PRIX TOTAL)
        $style = ($index === count($priceData) - 1) ? $totalRowStyle : null;
        $writer->addRow(WriterEntityFactory::createRowFromArray($row, $style));
    }

    // ======= FEUILLE 2 : BPU DÉTAILLÉ =======

    // Créer une nouvelle feuille pour le BPU détaillé
    $writer->addNewSheetAndMakeItCurrent();
    $writer->addRow(WriterEntityFactory::createRowFromArray(['BPU DÉTAILLÉ AVEC QUANTITÉS'], $headerStyle));
    $writer->addRow(WriterEntityFactory::createRowFromArray([''])); // Ligne vide

    // Lecture du BPU original
    $bpuData = readBPUFile($bpuFilePath);

    // Créer un mapping des descriptions aux quantités sélectionnées
    $quantiteParDescription = [];
    foreach ($data['recap'] as $item) {
        if (!empty($item['description']) && isset($item['quantity']) && $item['quantity'] > 0) {
            $quantiteParDescription[$item['description']] = $item['quantity'];
        }
    }

    // Parcourir les feuilles du BPU
    foreach ($bpuData as $sheetName => $rows) {
        // Ajouter le nom de la feuille comme titre
        $writer->addRow(WriterEntityFactory::createRowFromArray(['Feuille: ' . $sheetName], $headerStyle));

        // Vérifier s'il y a des données dans cette feuille
        if (empty($rows)) {
            $writer->addRow(WriterEntityFactory::createRowFromArray(['Aucune donnée dans cette feuille']));
            continue;
        }

        // En-têtes originaux du BPU
        $headers = $rows[0] ?? [];

        // Ajouter une colonne pour la quantité sélectionnée
        $headers[] = 'Quantité Sélectionnée';

        // Écrire la ligne d'en-tête
        $writer->addRow(WriterEntityFactory::createRowFromArray($headers, $headerStyle));

        // Trouver l'index de la colonne Description
        $descriptionIndex = -1;
        foreach ($headers as $i => $header) {
            if (stripos($header, 'designation') !== false ||
                stripos($header, 'description') !== false ||
                stripos($header, 'produit') !== false) {
                $descriptionIndex = $i;
                break;
            }
        }

        // Écrire les données
        for ($i = 1; $i < count($rows); $i++) {
            $row = $rows[$i];

            // Ajouter la quantité sélectionnée si la description correspond
            $quantite = '';
            if ($descriptionIndex >= 0 && isset($row[$descriptionIndex])) {
                $description = $row[$descriptionIndex];
                $quantite = $quantiteParDescription[$description] ?? '';
            }

            // Ajouter la valeur de la quantité à la fin
            $row[] = $quantite;

            // Écrire la ligne
            $writer->addRow(WriterEntityFactory::createRowFromArray($row));
        }

        // Ajouter une ligne vide après chaque feuille
        $writer->addRow(WriterEntityFactory::createRowFromArray(['']));
    }

    // Fermer le fichier
    $writer->close();

    // Envoyer le fichier au client
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment; filename="Export_Panneaux_' . date('Ymd_His') . '.xlsx"');
    header('Content-Length: ' . filesize($tempFilePath));
    header('Pragma: no-cache');
    header('Expires: 0');

    readfile($tempFilePath);

    // Supprimer le fichier temporaire
    @unlink($tempFilePath);

} catch (Exception $e) {
    // En cas d'erreur, renvoyer un message d'erreur
    header('HTTP/1.1 500 Internal Server Error');
    echo json_encode(['error' => 'Erreur lors de la génération du fichier Excel: ' . $e->getMessage()]);

    // Supprimer le fichier temporaire s'il existe
    if (file_exists($tempFilePath)) {
        @unlink($tempFilePath);
    }
}

/**
 * Lit le fichier BPU et retourne les données sous forme de tableau
 *
 * @param string $filePath Chemin vers le fichier BPU
 * @return array Données du fichier BPU
 */
function readBPUFile($filePath) {
    $result = [];

    try {
        // Création du lecteur pour fichier XLSX
        $reader = ReaderEntityFactory::createReaderFromFile($filePath);

        // Ouverture du fichier
        $reader->open($filePath);

        // Lecture des données
        foreach ($reader->getSheetIterator() as $sheet) {
            $sheetData = [];
            foreach ($sheet->getRowIterator() as $row) {
                // Conversion de la ligne en tableau de valeurs
                $sheetData[] = $row->toArray();
            }

            $result[$sheet->getName()] = $sheetData;
        }

        // Fermeture du lecteur
        $reader->close();

    } catch (Exception $e) {
        // En cas d'erreur, renvoyer un tableau vide
        error_log('Erreur lors de la lecture du fichier BPU: ' . $e->getMessage());
    }

    return $result;
}