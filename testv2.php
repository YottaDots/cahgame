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
$totalCardDeck = array();

foreach ($reader->getSheetIterator() as $sheet){
    $CardDeck = array();
    $setnr = 0;
    $setrowfound = 0;
    $iCell=0;
    $icarddeck=0;
    /** @var Row $row */
    foreach ($sheet->getRowIterator() as $iRow => $row){
        //the word "Set" is the start of a group of cards. It can happen that there are several "Sets" in one row, so we build an array for each sheet.
        //when teh sheet comes to it's end we go through the array and save it in the database.
        /*
         * Row1: Set, pack title,Special,versions (1 or more columns)
         * Row2: pack description
         * Row3: empty
         * Row4:cards
         *
         * array looks like:
         * [idset][columnr]
         * [idset][rownr]
         * [idset][maxcolumnnr]
         * [idset][packtitle]
         * [prompt][1][cardtitle]
         * [prompt][1][special]
         * [prompt][1][versions][v1.1]
         */
        $aRows = $row->getCells();
        $amountColmns =  count($aRows);
        $k = 0;
        foreach ( $aRows as $iColumn => $cell){
            $amountdecks = count($CardDeck);
            if($amountdecks > 0) {
                for ($c=0 ; $c < $amountdecks ; $c++) {
                  //for each "deck"we see if there is some information found
                   if($iColumn >= $CardDeck[$c]['columnr'] && isset($CardDeck[$c]['maxcolumnr']) && ($iColumn <= $CardDeck[$c]['columnr'] + $CardDeck[$c]['maxcolumnr'] ) && $iRow != $CardDeck[$c]['rownr'] ) {
                       if (!empty($cell->getValue()) ) {
                               switch (true) {
                                   case $iColumn == $CardDeck[$c]['columnr']+1 && $iRow == $CardDeck[$c]['rownr']+1:
                                       $CardDeck[$c]['carddeckdescription'] = $cell->getValue();
                                       break;
                                   case $iColumn == $CardDeck[$c]['columnr']:
                                       if (strtolower($cell->getValue()) == 'prompt') {
                                           $setcode = 'prompt';
                                       } else {
                                           $setcode = 'response';
                                       }
                                       break;
                                   case $iColumn == $CardDeck[$c]['columnr'] + 1:
                                       $CardDeck[$c][$setcode][$iRow]['cardtitle'] = $cell->getValue();
                                       break;
                                   case $iColumn == $CardDeck[$c]['columnr'] + 2:
                                       $CardDeck[$c][$setcode][$iRow]['special'] = $cell->getValue();
                                       break;
                                   case  $iColumn > $CardDeck[$c]['columnr'] + 2:
                                       if (!empty($cell->getValue())) {
                                           $CardDeck[$c][$setcode][$iRow]['versions'][] = $cell->getValue();
                                       }
                                       break;
                               }
                           //$CardDeck[$c]['versions'][] = $cell->getValue();
                       }
                   }
                }
            }
            if($cell->getValue() == 'Set'){
                $setrowfound = 1;
                $icarddeck = 0;
                $iCell++;
                $CardDeck[$setnr]['columnr'] = $iColumn;
                $CardDeck[$setnr]['rownr'] = $iRow;
                $CardDeck[$setnr]['getValue'] = $cell->getValue();
            }
            if ($setrowfound == 1) {
                switch(true) {
                    case $iColumn == $CardDeck[$setnr]['columnr'] + 1:
                        $CardDeck[$setnr]['carddecktitle'] = $cell->getValue();
                        break;
                    case $iColumn == $CardDeck[$setnr]['columnr'] + 2:
                        $CardDeck[$setnr]['special'] = $cell->getValue();
                        break;
                    case  $iColumn > $CardDeck[$setnr]['columnr'] + 2:
                        if (!empty($cell->getValue())) {
                            $CardDeck[$setnr]['versions'][] = $cell->getValue();
                        }
                        break;
                }
            }
            if((empty($cell->getValue()) && $setrowfound == 1) OR ( $k+1 == $amountColmns && isset($CardDeck[$setnr]))){
                if($k+1 == $amountColmns) {
                    $CardDeck[$setnr]['maxcolumnr'] = $amountColmns - $CardDeck[$setnr]['columnr'];
                } else {
                    $CardDeck[$setnr]['maxcolumnr'] = $iColumn - $CardDeck[$setnr]['columnr'];
                }
                $setrowfound = 0;
                $iCell=0;
                $icarddeck=0;
                $setnr++;
            }
            $k++;
        }
    }
    $totalCardDeck[$sheet->getName()] = $CardDeck;
}
print_r($totalCardDeck);
$reader->close();
