<?php
use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
class FileReader
{
    private $filePath;
    private $reader;

    public function __construct($filePath)
    {
        $this->filePath = $filePath;
        $this->reader = ReaderEntityFactory::createReaderFromFile($this->filePath);
    }

    public function read()
    {
        $this->reader->open($this->filePath);

        foreach ($this->reader->getSheetIterator() as $sheet) {
            foreach ($sheet->getRowIterator() as $row) {
                // Process each row
                $cells = $row->getCells();
            }
        }

        $this->reader->close();
    }
}