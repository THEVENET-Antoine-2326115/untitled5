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

    // Style commun pour toutes les lignes
    $baseRowStyle = (new StyleBuilder())
        ->build();

    // Styles pour les lignes alternées - uniquement la couleur change
    $rowStyleEven = (new StyleBuilder())
        ->setBackgroundColor('FFFFFF') // Blanc (hex)
        ->build();

    $rowStyleOdd = (new StyleBuilder())
        ->setBackgroundColor('F0F0F0') // Gris très clair (hex)
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
        'Pièce associée',
        'Transport Barreaux',
        'Manutention Barreaux',
        'Nbr Demi-Journées',
        'Grutage Barreaux',
        'Transport Panneaux',
        'Manutention Panneaux',
        'Grutage Panneaux',
        'Emballage',

    ];
    $writer->addRow(WriterEntityFactory::createRowFromArray($recapHeaders, $headerStyle));

    // Données du tableau récapitulatif
    foreach ($data['recap'] as $item) {
        // Convertir toutes les valeurs en chaînes
        $rowData = [
            (string)($item['site'] ?? ''),
            (string)($item['description'] ?? ''),
            (string)($item['quantity'] ?? ''),
            (string)($item['pieceAssociee'] ?? ''),
            (string)($item['transportBarreaux'] ?? ''),
            (string)($item['manutentionBarreaux'] ?? ''),
            (string)($item['manutentionDays'] ?? ''),
            (string)($item['dechargementBarreaux'] ?? ''),
            (string)($item['transportPanneaux'] ?? ''),
            (string)($item['manutentionPanneaux'] ?? ''),
            (string)($item['grutagePanneaux'] ?? ''),
            (string)($item['emballage'] ?? ''),

        ];
        $writer->addRow(WriterEntityFactory::createRowFromArray($rowData));
    }

    // Ajouter un espace
    $writer->addRow(WriterEntityFactory::createRowFromArray(['']));
    $writer->addRow(WriterEntityFactory::createRowFromArray(['']));

    // ======= RÉCAPITULATIF DES PRIX DÉTAILLÉ =======

    // Ajouter le récapitulatif des prix avec des colonnes détaillées
    $writer->addRow(WriterEntityFactory::createRowFromArray(['RÉCAPITULATIF DÉTAILLÉ DES PRIX'], $headerStyle));
    $writer->addRow(WriterEntityFactory::createRowFromArray([''])); // Ligne vide

    // En-têtes du tableau de prix détaillé - Structure avec colonnes distinctes pour quantité, camions et demi-journées
    $priceHeaders = [
        'Désignation',
        'Quantité',
        'Nombre camion',
        'Nombre Demi-journée',
        'Prix unitaire hors taxe (€)',
        'Formule de calcul',
        'Montant total hors taxe (€)'
    ];
    $writer->addRow(WriterEntityFactory::createRowFromArray($priceHeaders, $headerStyle));

    // Données détaillées des prix
    $priceData = [];

    // 1. Forfait logistique
    if (isset($data['prices']['forfaitLogistique'])) {
        $prixLogistique = $data['prices']['forfaitLogistique'];
        $priceData[] = [
            'Forfait logistique',
            $data['prices']['totalPanneaux'] ?? '0',
            '',
            '',
            $prixLogistique['prixUnitaire'] ?? '0,00',
            $data['prices']['totalPanneaux'] . ' × ' . $prixLogistique['prixUnitaire'] . ' €',
            number_format((float)($prixLogistique['montant'] ?? 0), 2, ',', ' ')
        ];
    }

    // 2. Forfait suivi
    if (isset($data['prices']['forfaitSuivi'])) {
        $prixSuivi = $data['prices']['forfaitSuivi'];
        $priceData[] = [
            'Forfait suivi',
            $data['prices']['totalPanneaux'] ?? '0',
            '',
            '',
            $prixSuivi['prixUnitaire'] ?? '0,00',
            $data['prices']['totalPanneaux'] . ' × ' . $prixSuivi['prixUnitaire'] . ' €',
            number_format((float)($prixSuivi['montant'] ?? 0), 2, ',', ' ')
        ];
    }

    // 3. Soudage des barreaux
    $priceData[] = [
        'Soudage des barreaux',
        $data['prices']['nbPanneauxSupStVallier'] ?? '0',
        '',
        '',
        '533,56',
        $data['prices']['nbPanneauxSupStVallier'] . ' × 533,56 €',
        number_format((float)($data['prices']['prixSoudage'] ?? 0), 2, ',', ' ')
    ];

    // 4. Transport des barreaux
    if (isset($data['prices']['transportBarreaux']) && !empty($data['prices']['transportBarreaux']['detailParSite'])) {
        foreach ($data['prices']['transportBarreaux']['detailParSite'] as $site => $details) {
            $formule = "Poids total: " . number_format($details['poidsTotal'] ?? 0, 4, ',', ' ') . " t / 19 t = " .
                $details['nombreCamions'] . " camions × " . number_format($details['prixUnitaire'], 2, ',', ' ') . " €";

            $priceData[] = [
                'Transport des barreaux - ' . $site,
                '',
                $details['nombreCamions'],
                '',
                number_format($details['prixUnitaire'] ?? 0, 2, ',', ' '),
                $formule,
                number_format((float)($details['prixTransport'] ?? 0), 2, ',', ' ')
            ];
        }
    }

    // 5. Transport des panneaux
    if (isset($data['prices']['transportPanneaux']) && !empty($data['prices']['transportPanneaux']['detailParSite'])) {
        foreach ($data['prices']['transportPanneaux']['detailParSite'] as $site => $details) {
            $formule = "Quantité: " . ($details['quantiteTotale'] ?? 0) . " panneaux / 5 = " .
                $details['nombreCamions'] . " camions × " . number_format($details['prixUnitaire'], 2, ',', ' ') . " €";

            $priceData[] = [
                'Transport des panneaux - ' . $site,
                $details['quantiteTotale'] ?? '0',
                $details['nombreCamions'],
                '',
                number_format($details['prixUnitaire'] ?? 0, 2, ',', ' '),
                $formule,
                number_format((float)($details['prixTransport'] ?? 0), 2, ',', ' ')
            ];
        }
    }

    // 6. Manutention des barreaux
    if (isset($data['prices']['manutention']) && !empty($data['prices']['manutention']['detailParSite'])) {
        foreach ($data['prices']['manutention']['detailParSite'] as $site => $details) {
            $formule = $details['nbDemiJournees'] . " demi-journées × " . number_format($details['prixUnitaire'], 2, ',', ' ') . " €";

            $priceData[] = [
                'Manutention des barreaux - ' . $site,
                '',
                '',
                $details['nbDemiJournees'],
                number_format($details['prixUnitaire'] ?? 0, 2, ',', ' '),
                $formule,
                number_format((float)($details['prixManutention'] ?? 0), 2, ',', ' ')
            ];
        }
    }

    // 7. Manutention des panneaux
    if (isset($data['prices']['manutentionPanneaux']) && !empty($data['prices']['manutentionPanneaux']['detailParSite'])) {
        foreach ($data['prices']['manutentionPanneaux']['detailParSite'] as $site => $details) {
            $formule = "Quantité: " . ($details['quantiteTotale'] ?? 0) . " panneaux / 5 = " .
                $details['nombreCamions'] . " demi-journées × " . number_format($details['prixUnitaire'], 2, ',', ' ') . " €";

            $priceData[] = [
                'Manutention des panneaux - ' . $site,
                $details['quantiteTotale'] ?? '0',
                '',
                $details['nombreCamions'],
                number_format($details['prixUnitaire'] ?? 0, 2, ',', ' '),
                $formule,
                number_format((float)($details['prixManutention'] ?? 0), 2, ',', ' ')
            ];
        }
    }

    // 8. Grutage des barreaux
    if (isset($data['prices']['grutage']) && !empty($data['prices']['grutage']['detailParSite'])) {
        foreach ($data['prices']['grutage']['detailParSite'] as $site => $details) {
            $formule = $details['nbDemiJournees'] . " demi-journées × 882,00 €";

            $priceData[] = [
                'Grutage des barreaux - ' . $site,
                '',
                '',
                $details['nbDemiJournees'],
                '882,00',
                $formule,
                number_format((float)($details['prixGrutage'] ?? 0), 2, ',', ' ')
            ];
        }
    }

    // 9. Grutage des panneaux
    if (isset($data['prices']['grutagePanneaux']) && !empty($data['prices']['grutagePanneaux']['detailParSite'])) {
        foreach ($data['prices']['grutagePanneaux']['detailParSite'] as $site => $details) {
            $formule = $details['nombreCamions'] . " demi-journées × 882,00 €";

            $priceData[] = [
                'Grutage des panneaux - ' . $site,
                '',
                '',
                $details['nombreCamions'],
                '882,00',
                $formule,
                number_format((float)($details['prixGrutage'] ?? 0), 2, ',', ' ')
            ];
        }
    }

    // 10. Emballage
    $priceData[] = [
        'Emballage',
        $data['prices']['totalPanneaux'] ?? '0',
        '',
        '',
        '235,94',
        $data['prices']['totalPanneaux'] . ' × 235,94 €',
        number_format((float)($data['prices']['prixEmballage'] ?? 0), 2, ',', ' ')
    ];

    // 11. Montage des panneaux
    if (isset($data['prices']['detailMontage']) && !empty($data['prices']['detailMontage'])) {
        foreach ($data['prices']['detailMontage'] as $montage) {
            $priceData[] = [
                'Montage - ' . ($montage['description'] ?? 'Panneau') . ' - ' . ($montage['site'] ?? ''),
                $montage['quantity'] ?? '0',
                '',
                '',
                number_format($montage['prixUnitaire'] ?? 0, 2, ',', ' '),
                $montage['quantity'] . ' × ' . number_format($montage['prixUnitaire'], 2, ',', ' ') . ' €',
                number_format((float)($montage['sousTotal'] ?? 0), 2, ',', ' ')
            ];
        }
    } else {
        $priceData[] = [
            'Montage des panneaux',
            '',
            '',
            '',
            '',
            'Somme des sous-totaux de montage',
            number_format((float)($data['prices']['prixMontage'] ?? 0), 2, ',', ' ')
        ];
    }

    // 12. PRIX TOTAL
    $priceData[] = [
        'PRIX TOTAL',
        '',
        '',
        '',
        '',
        'Somme de tous les montants',
        number_format((float)($data['prices']['prixTotal'] ?? 0), 2, ',', ' ')
    ];

    // Ajouter chaque ligne avec couleurs alternées
    foreach ($priceData as $index => $row) {
        // Utiliser le style total pour la dernière ligne (PRIX TOTAL)
        if ($index === count($priceData) - 1) {
            $style = $totalRowStyle;
        }

        // Uniformiser les lignes en appliquant le texte de la même façon
        // Convertir chaque valeur en chaîne explicitement
        $processedRow = [];
        foreach ($row as $cellValue) {
            // Assurer que toutes les valeurs sont des chaînes
            $processedRow[] = (string)$cellValue;
        }

        $writer->addRow(WriterEntityFactory::createRowFromArray($processedRow, $style));
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

    // Créer un mapping des descriptions aux quantités sélectionnées en fonction du site
    $quantiteParSiteEtDescription = [];
    foreach ($data['recap'] as $item) {
        if (!empty($item['description']) && isset($item['quantity']) && $item['quantity'] > 0 && !isset($item['isTransportOnly'])) {
            // Utiliser la combinaison site + description comme clé
            $key = ($item['site'] ?? '') . '|' . $item['description'];

            // Si cette combinaison existe déjà, additionner les quantités
            if (isset($quantiteParSiteEtDescription[$key])) {
                $quantiteParSiteEtDescription[$key] += (int)$item['quantity'];
            } else {
                $quantiteParSiteEtDescription[$key] = (int)$item['quantity'];
            }

            error_log("Combinaison site+description ajoutée: '$key' => " . $quantiteParSiteEtDescription[$key]);
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

        // Ajouter une colonne pour la quantité sélectionnée en position P (index 15)
        // Assurer qu'il y a suffisamment de colonnes jusqu'à P
        while (count($headerStrings) < 15) {
            $headerStrings[] = ''; // Ajouter des colonnes vides si nécessaire
        }
        $headerStrings[15] = 'Quantité Sélectionnée'; // Placer l'en-tête en colonne P

        // Écrire la ligne d'en-tête
        $writer->addRow(WriterEntityFactory::createRowFromArray($headerStrings, $headerStyle));

        // Colonne B (index 1) est la colonne de description
        $descriptionIndex = 1;

        // Variable pour suivre le site courant
        $currentSite = '';

        // Écrire les données avec couleurs alternées
        for ($i = 1; $i < count($rows); $i++) {
            $row = $rows[$i];
            $rowData = [];

            // Mettre à jour le site courant si la colonne A n'est pas vide
            if (!empty($row[0])) {
                $potentialSite = trim((string)$row[0]);
                // Vérifier si c'est un site (un mot sans chiffres ni espaces)
                if (preg_match('/^[A-Za-z\-]+$/', $potentialSite)) {
                    $currentSite = $potentialSite;
                    error_log("Site courant mis à jour: $currentSite");
                }
            }

            // Convertir chaque cellule en chaîne
            foreach ($row as $cell) {
                $rowData[] = (string)$cell;
            }

            // S'assurer que nous avons suffisamment de colonnes jusqu'à P (index 15)
            while (count($rowData) < 15) {
                $rowData[] = ''; // Ajouter des colonnes vides si nécessaire
            }

            // Ajouter la quantité sélectionnée si la description correspond
            $quantite = '';
            if (isset($row[$descriptionIndex]) && !empty($currentSite)) {
                $description = (string)$row[$descriptionIndex];

                // Créer la clé combinée site+description
                $key = $currentSite . '|' . $description;

                if (isset($quantiteParSiteEtDescription[$key])) {
                    $quantite = (string)$quantiteParSiteEtDescription[$key];
                    error_log("Correspondance trouvée pour '$key': Quantité: $quantite");
                }
            }

            // Placer la valeur de la quantité en colonne P (index 15)
            $rowData[15] = $quantite;

            // Appliquer un style alterné aux lignes (sauf en-têtes)
            $style = ($i % 2 === 0) ? $rowStyleEven : $rowStyleOdd;

            // Uniformiser les lignes en appliquant le texte de la même façon
            $processedRow = [];
            foreach ($rowData as $cellValue) {
                // Assurer que toutes les valeurs sont des chaînes
                $processedRow[] = (string)$cellValue;
            }

            // Écrire la ligne
            $writer->addRow(WriterEntityFactory::createRowFromArray($processedRow, $style));
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
?>