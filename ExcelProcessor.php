<?php
// ExcelProcessor.php
require_once 'ExcelReader.php';

class ExcelProcessor
{
    private $excelReader;
    private $data = [];
    private $productTotals = [];
    private $filePath;

    public function __construct($filePath)
    {
        $this->filePath = $filePath;
        $this->excelReader = new ExcelReader($filePath);
    }

    public function processData()
    {
        // Lecture du fichier Excel
        if (!$this->excelReader->read()) {
            return false;
        }

        // Récupération des données
        $this->data = $this->excelReader->getData();

        // Déterminer le type de fichier
        $isWeightFile = strpos($this->filePath, 'ref') !== false;

        // Pour le fichier ref.xlsm, nous n'avons pas besoin de calculer les totaux
        if ($isWeightFile) {
            return true;
        }

        // Calcul des totaux par produit uniquement pour BPU_vrai.xlsx
        foreach ($this->data as $sheetName => $rows) {
            if (count($rows) <= 0) {
                continue; // Ignorer les feuilles vides
            }

            $headers = $rows[0];

            $prixCol = $this->findColumn($headers, ['prix', 'price']);
            $qteCol = $this->findColumn($headers, ['quantité', 'quantity', 'qté', 'qty']);
            $produitCol = $this->findColumn($headers, ['produit', 'product', 'article', 'item', 'designation', 'description']);
            $livraisonCol = $this->findColumn($headers, ['livraison', 'shipping']);

            // Si on ne trouve pas les colonnes nécessaires, on passe à la feuille suivante
            if ($prixCol === false || $qteCol === false) {
                continue;
            }

            $this->productTotals[$sheetName] = [
                'products' => [],
                'rows' => []
            ];

            // Parcourir les lignes de données
            for ($i = 1; $i < count($rows); $i++) {
                $prix = $this->toNumber($rows[$i][$prixCol] ?? 0);
                $qte = $this->toNumber($rows[$i][$qteCol] ?? 0);
                $livraison = ($livraisonCol !== false) ? $this->toNumber($rows[$i][$livraisonCol] ?? 0) : 0;
                $produit = ($produitCol !== false) ? ($rows[$i][$produitCol] ?? "Produit #$i") : "Produit #$i";

                $total = ($prix * $qte) + $livraison;
                $this->productTotals[$sheetName]['rows'][$i] = $total;

                // Stockage du prix total par produit
                if (!isset($this->productTotals[$sheetName]['products'][$produit])) {
                    $this->productTotals[$sheetName]['products'][$produit] = $total;
                } else {
                    // Si le produit existe déjà, on ajoute au total
                    $this->productTotals[$sheetName]['products'][$produit] += $total;
                }
            }
        }

        return true;
    }

    private function findColumn($headers, $possibleNames)
    {
        foreach ($headers as $index => $header) {
            if (!$header || !is_string($header)) {
                continue;
            }

            $header = strtolower($header);
            foreach ($possibleNames as $name) {
                if ($header === strtolower($name) || strpos($header, strtolower($name)) !== false) {
                    return $index;
                }
            }
        }
        return false;
    }

    private function toNumber($value)
    {
        if (is_numeric($value)) {
            return floatval($value);
        }

        if (is_string($value)) {
            $value = str_replace(',', '.', $value);
            $value = preg_replace('/[^0-9.]/', '', $value);
            return is_numeric($value) ? floatval($value) : 0;
        }

        return 0;
    }

    public function getData()
    {
        return $this->data;
    }

    public function getProductTotals()
    {
        return $this->productTotals;
    }

    public function getFilePath()
    {
        return $this->filePath;
    }
}