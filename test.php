<?php
/*
 * we use spout based upon the review on https://www.nidup.io/blog/manipulate-excel-files-in-php
 * https://github.com/box/spout
 */
require_once 'spout-master/src/Spout/Autoloader/autoload.php';

use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
//$filePath = 'Cards Against Humanity.xlsx';
$filePath = 'test.xlsx';
$reader = ReaderEntityFactory::createXLSXReader();

$reader->open($filePath);

//foreach ($reader->getSheetIterator() as $sheet) {
//    echo $sheet->getName().' '. $sheet->isVisible().' '. $sheet->isActive().PHP_EOL;
//    $i = 0;
//    foreach ($sheet->getRowIterator() as $row) {
//        $i++;
//        // do stuff with the row
//        $aRow = $row->getCells();
//        foreach ($aRow as $cell){
//            $cell_value = $cell->getValue();
//            echo "Cell Value = $cell_value<br>\n";
//        }
////        print_r($aRow);
////        print_r($row->getStyle());
////        print_r($row->getStyle());
//
////        $color = $row->getStyle();
////         echo $color;
//    }
//    echo '$i:.'.$i;
//}
/** @var Sheet $sheet */
foreach ($reader->getSheetIterator() as $sheet){
    echo sprintf('Sheet: %s', $sheet->getName()) . PHP_EOL;

    /** @var Row $row */
    foreach ($sheet->getRowIterator() as $index => $row){
        echo sprintf('Row: %d', $index) . PHP_EOL;
        foreach ($row->getCells() as $indexCell => $cell){
            echo
                sprintf(
                    'Cell index: %d, value: %s, background color: %s, font color: %s',
                    $indexCell,
                    $cell->getValue(),
                    $cell->getStyle()->getBackgroundColor(),
                    $cell->getStyle()->getFontColor()
                ) .
                PHP_EOL
            ;
        }
    }
}
$reader->close();
