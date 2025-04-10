<?php
// ExcelReader.php
require 'vendor/autoload.php';

use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;

class ExcelReader
{
    private $filePath;
    private $data = [];

    public function __construct($filePath)
    {
        $this->filePath = $filePath;
    }

    public function read()
    {
        try {
            // Création du lecteur pour fichier XLSX
            $reader = ReaderEntityFactory::createReaderFromFile($this->filePath);

            // Ouverture du fichier
            $reader->open($this->filePath);

            // Lecture des données
            foreach ($reader->getSheetIterator() as $sheet) {
                $sheetData = [];
                foreach ($sheet->getRowIterator() as $row) {
                    // Conversion de la ligne en tableau de valeurs
                    $sheetData[] = $row->toArray();
                }
                $this->data[$sheet->getName()] = $sheetData;
            }

            // Fermeture du lecteur
            $reader->close();

            return true;
        } catch (Exception $e) {
            echo 'Erreur lors de la lecture du fichier Excel: ' . $e->getMessage();
            return false;
        }
    }

    public function getData()
    {
        return $this->data;
    }
}