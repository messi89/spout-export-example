<?php
//ini_set('display_errors', 1);
//ini_set('display_startup_errors', 1);
//error_reporting(E_ALL);

require 'vendor/autoload.php';

use Box\Spout\Common\Entity\Style\CellAlignment;
use Box\Spout\Common\Entity\Style\Color;
use Box\Spout\Writer\Common\Creator\Style\StyleBuilder;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;

/**
 * @param int $length
 * @return string
 */
function generateRandomString($length = 10) : string {
    $characters = '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ';
    $charactersLength = strlen($characters);
    $randomString = '';
    for ($i = 0; $i < $length; $i++) {
        $randomString .= $characters[rand(0, $charactersLength - 1)];
    }
    return $randomString;
}

/**
 * @param int $size
 * @return array
 */
function generateRandomArray($size= 500) : array {
    $index = 1;
    $data = array();

    while($index <= $size) {
        array_push($data, [
            'id' => $index,
            'name' => generateRandomString(),
            'username' => generateRandomString(6),
            'email' => generateRandomString(8).'@gmail.com',
            'address' => generateRandomString(),
            'region' => generateRandomString(4),
            'date' => date(DATE_ISO8601)
        ]);
        $index++;
    }

    return $data;
}

/**
 * @param $dataArray
 * @param string $fileName
 * @param null $headerArray
 * @throws \Box\Spout\Common\Exception\IOException
 * @throws \Box\Spout\Common\Exception\InvalidArgumentException
 * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
 */
function generateExcelFileFromArray($dataArray, $fileName = 'export', $headerArray = null) : void {
    $writer = WriterEntityFactory::createXLSXWriter();

    $writer->openToBrowser($fileName.'.xlsx'); //->openToFile(getcwd ().'/export.xlsx')

    // header styling
    $style = (new StyleBuilder())
        ->setFontBold()
        ->setFontSize(12)
        ->setFontColor(Color::BLACK)
        ->setShouldWrapText(false)
        ->setCellAlignment(CellAlignment::LEFT)
        ->setBackgroundColor(Color::rgb(232,232,232))
        ->build();

    // header
    $header = WriterEntityFactory::createRowFromArray($headerArray !== null ? $headerArray : array_keys($dataArray[0]), $style);
    $writer->addRow($header);

    //stream data to writer
    foreach ($dataArray as $line) {
        $rowFromValues = WriterEntityFactory::createRowFromArray($line);
        $writer->addRow($rowFromValues);
    }

    $writer->close();
}

try {
    //generate data
    $data = generateRandomArray(5000);

    //generate excel file
    generateExcelFileFromArray($data);
}
catch(Exception $e) {
    echo $e->getMessage();
}
