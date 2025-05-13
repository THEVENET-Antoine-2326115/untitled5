<?php
// Supprimez l'affichage des messages de dépréciation mais conservez les autres erreurs
error_reporting(E_ALL & ~E_DEPRECATED);

// Utilisez le tampon de sortie pour éviter les problèmes d'en-têtes
ob_start();

// Inclure les dépendances
require 'vendor/autoload.php';

// Utiliser les classes Spout pour la manipulation des fichiers Excel
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Writer\Common\Creator\Style\StyleBuilder;
use Box\Spout\Common\Entity\Row;
use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;

// Journaliser l'erreur pour débogage (facultatif)
error_log("Début de l'export Excel: " . date('Y-m-d H:i:s'));

// Définir le chemin du fichier temporaire dans une portée plus large
$tempFilePath = "temp_export_" . uniqid() . ".xlsx";

try {
    // Récupérer les données JSON envoyées par le client
    $jsonData = file_get_contents('php://input');
    $data = json_decode($jsonData, true);

    // Vérifier que les données ont été décodées correctement
    if (json_last_error() !== JSON_ERROR_NONE) {
        throw new Exception("Erreur de décodage JSON: " . json_last_error_msg());
    }

    // Vérifier si les données sont valides
    if (!$data || !isset($data['recap']) || !isset($data['prices'])) {
        throw new Exception("Données invalides ou incomplètes");
    }

    // Chemins des fichiers
    $bpuFilePath = "BPU_vrai.xlsx";
    if (!file_exists($bpuFilePath)) {
        throw new Exception("Le fichier BPU n'existe pas: $bpuFilePath");
    }

    // Créer un nouveau fichier Excel
    $writer = WriterEntityFactory::createXLSXWriter();
    $writer->openToFile($tempFilePath);

    // Créer des styles pour les en-têtes et les cellules
    $headerStyle = (new StyleBuilder())
        ->setFontBold()
        ->setFontSize(12)
        ->setBackgroundColor('ADD8E6') // Bleu clair (hex)
        ->setFontColor('000000')       // Noir (hex)
        ->build();

    $totalRowStyle = (new StyleBuilder())
        ->setFontBold()
        ->setFontSize(12)
        ->setBackgroundColor('DDDDDD') // Gris clair (hex)
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
        // Convertir toutes les valeurs en chaînes
        $rowData = [
            (string)($item['site'] ?? ''),
            (string)($item['description'] ?? ''),
            (string)($item['quantity'] ?? ''),
            (string)($item['transportBarreaux'] ?? ''),
            (string)($item['manutentionBarreaux'] ?? ''),
            (string)($item['manutentionDays'] ?? ''),
            (string)($item['dechargementBarreaux'] ?? ''),
            (string)($item['transportPanneaux'] ?? ''),
            (string)($item['manutentionPanneaux'] ?? ''),
            (string)($item['grutagePanneaux'] ?? ''),
            (string)($item['emballage'] ?? ''),
            isset($item['totalPrice']) ? number_format((float)$item['totalPrice'], 2, ',', ' ') . ' €' : ''
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
        ['Forfaits (logistique et suivi)', number_format((float)($data['prices']['prixForfaits'] ?? 0), 2, ',', ' ') . ' €'],
        ['Soudage des barreaux', number_format((float)($data['prices']['prixSoudage'] ?? 0), 2, ',', ' ') . ' €'],
        ['Transport des barreaux', number_format((float)($data['prices']['prixTransportBarreaux'] ?? 0), 2, ',', ' ') . ' €'],
        ['Transport des panneaux', number_format((float)($data['prices']['prixTransportPanneaux'] ?? 0), 2, ',', ' ') . ' €'],
        ['Manutention des barreaux', number_format((float)($data['prices']['prixManutention'] ?? 0), 2, ',', ' ') . ' €'],
        ['Manutention des panneaux', number_format((float)($data['prices']['prixManutentionPanneaux'] ?? 0), 2, ',', ' ') . ' €'],
        ['Grutage des barreaux', number_format((float)($data['prices']['prixGrutage'] ?? 0), 2, ',', ' ') . ' €'],
        ['Grutage des panneaux', number_format((float)($data['prices']['prixGrutagePanneaux'] ?? 0), 2, ',', ' ') . ' €'],
        ['Emballage', number_format((float)($data['prices']['prixEmballage'] ?? 0), 2, ',', ' ') . ' €'],
        ['Montage des panneaux', number_format((float)($data['prices']['prixMontage'] ?? 0), 2, ',', ' ') . ' €'],
        ['PRIX TOTAL', number_format((float)($data['prices']['prixTotal'] ?? 0), 2, ',', ' ') . ' €']
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

    try {
        // Lecture du BPU original
        $bpuData = readBPUFile($bpuFilePath);
    } catch (Exception $e) {
        // En cas d'échec, créer une structure minimale
        error_log("Erreur lors de la lecture du BPU original: " . $e->getMessage());
        $bpuData = ['Feuille1' => [['Pas de données disponibles']]];
    }

    // Créer un mapping des descriptions aux quantités sélectionnées (en normalisant les descriptions)
    $quantiteParDescription = [];
    foreach ($data['recap'] as $item) {
        if (!empty($item['description']) && isset($item['quantity']) && $item['quantity'] > 0 && !isset($item['isTransportOnly'])) {
            // Normaliser la description pour la recherche (enlever les espaces supplémentaires, mettre en minuscules)
            $descNormalisee = normaliserTexte($item['description']);
            $quantiteParDescription[$descNormalisee] = $item['quantity'];

            // Ajout de log pour débogage
            error_log("Description ajoutée au mapping: '$descNormalisee' => " . $item['quantity']);
        }
    }

    // Parcourir les feuilles du BPU
    foreach ($bpuData as $sheetName => $rows) {
        // Ajouter le nom de la feuille comme titre
        $writer->addRow(WriterEntityFactory::createRowFromArray(['Feuille: ' . (string)$sheetName], $headerStyle));

        // Vérifier s'il y a des données dans cette feuille
        if (empty($rows)) {
            $writer->addRow(WriterEntityFactory::createRowFromArray(['Aucune donnée dans cette feuille']));
            continue;
        }

        // En-têtes originaux du BPU
        $headers = $rows[0] ?? [];
        $headerStrings = [];

        // Convertir tous les en-têtes en chaînes
        foreach ($headers as $header) {
            $headerStrings[] = (string)$header;
        }

        // Ajouter une colonne pour la quantité sélectionnée
        $headerStrings[] = 'Quantité Sélectionnée';

        // Écrire la ligne d'en-tête
        $writer->addRow(WriterEntityFactory::createRowFromArray($headerStrings, $headerStyle));

        // Trouver l'index de la colonne Description
        $descriptionIndex = -1;
        foreach ($headerStrings as $i => $header) {
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
            $rowData = [];

            // Convertir chaque cellule en chaîne
            foreach ($row as $cell) {
                $rowData[] = (string)$cell;
            }

            // Ajouter la quantité sélectionnée si la description correspond
            $quantite = '';
            if ($descriptionIndex >= 0 && isset($row[$descriptionIndex])) {
                $description = (string)$row[$descriptionIndex];
                $descNormalisee = normaliserTexte($description);

                // Log pour débogage
                // error_log("Recherche de correspondance pour: '$descNormalisee'");

                if (isset($quantiteParDescription[$descNormalisee])) {
                    $quantite = (string)$quantiteParDescription[$descNormalisee];
                    error_log("Correspondance trouvée! Quantité: $quantite pour description: $description");
                } else {
                    // Recherche plus souple - vérifier si la description normalisée contient ou est contenue dans une des clés
                    foreach ($quantiteParDescription as $key => $qty) {
                        if (stripos($descNormalisee, $key) !== false || stripos($key, $descNormalisee) !== false) {
                            $quantite = (string)$qty;
                            error_log("Correspondance partielle trouvée! '$descNormalisee' ~ '$key', Quantité: $quantite");
                            break;
                        }
                    }
                }
            }

            // Ajouter la valeur de la quantité à la fin
            $rowData[] = $quantite;

            // Écrire la ligne
            $writer->addRow(WriterEntityFactory::createRowFromArray($rowData));
        }

        // Ajouter une ligne vide après chaque feuille
        $writer->addRow(WriterEntityFactory::createRowFromArray(['']));
    }

    // Fermer le fichier
    $writer->close();

    // Vider le tampon de sortie pour éviter les problèmes d'en-têtes
    ob_end_clean();

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
    // En cas d'erreur, enregistrer dans les logs et renvoyer un message d'erreur
    error_log("Erreur lors de l'export Excel: " . $e->getMessage() . "\n" . $e->getTraceAsString());

    // Vider le tampon de sortie
    ob_end_clean();

    // Envoyer un message d'erreur au client
    header('HTTP/1.1 500 Internal Server Error');
    header('Content-Type: application/json');
    echo json_encode(['error' => 'Erreur lors de la génération du fichier Excel: ' . $e->getMessage()]);

    // Supprimer le fichier temporaire s'il existe
    if (isset($tempFilePath) && file_exists($tempFilePath)) {
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
                // Conversion de la ligne en tableau de valeurs avec conversion sécurisée
                $rowArray = [];
                foreach ($row->toArray() as $cellValue) {
                    // Traitement sécurisé des valeurs de cellule
                    $processedValue = $cellValue;

                    // Convertir les objets en chaînes selon leur type
                    if (is_object($processedValue)) {
                        if (method_exists($processedValue, 'format') && $processedValue instanceof \DateTime) {
                            $processedValue = $processedValue->format('d/m/Y');
                        } elseif (method_exists($processedValue, '__toString')) {
                            $processedValue = (string)$processedValue;
                        } else {
                            $processedValue = 'Objet: ' . get_class($processedValue);
                        }
                    } elseif (is_array($processedValue)) {
                        $processedValue = 'Tableau de données';
                    } elseif ($processedValue === null) {
                        $processedValue = '';
                    }

                    $rowArray[] = $processedValue;
                }

                $sheetData[] = $rowArray;
            }

            $result[$sheet->getName()] = $sheetData;
        }

        // Fermeture du lecteur
        $reader->close();

    } catch (Exception $e) {
        // En cas d'erreur, enregistrer dans les logs
        error_log('Erreur lors de la lecture du fichier BPU: ' . $e->getMessage() . "\n" . $e->getTraceAsString());
        throw $e;
    }

    return $result;
}

/**
 * Normalise un texte pour faciliter les comparaisons
 * (retire les espaces supplémentaires, convertit en minuscules, supprime les accents)
 *
 * @param string $texte Texte à normaliser
 * @return string Texte normalisé
 */
function normaliserTexte($texte) {
    // Convertir en minuscules et supprimer les espaces supplémentaires
    $texte = mb_strtolower(trim($texte));
    $texte = preg_replace('/\s+/', ' ', $texte);

    // Remplacer les accents (version simple)
    $recherche = ['é', 'è', 'ê', 'ë', 'à', 'â', 'ä', 'î', 'ï', 'ô', 'ö', 'ù', 'û', 'ü', 'ç', 'ñ'];
    $remplace = ['e', 'e', 'e', 'e', 'a', 'a', 'a', 'i', 'i', 'o', 'o', 'u', 'u', 'u', 'c', 'n'];
    $texte = str_replace($recherche, $remplace, $texte);

    return $texte;
}
?>