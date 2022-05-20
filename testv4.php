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
//error: currently we add all the rows to the different sets which are in the same collumn
//maybe we should save the coordinates of each set and when we colline the we need to stop adding to that set.

/*
 * data groter dan start en kleiner dan eind
 *
 */

foreach ($reader->getSheetIterator() as $sheet){
    $CardDeck = array();
    $setnr = 0;
    $highestrow = 0;
    $setrowfound = 0;
    $rowbreakers = array();
    $iCell=0;
    $icarddeck=0;
    /** @var Row $row */
    foreach ($sheet->getRowIterator() as $iRow => $row){
        //the word "Set" is the start of a group of cards. It can happen that there are several "Sets" in one row, so we build an array for each sheet.
        //when teh sheet comes to it's end we go through the array and save it in the database.
        if( $highestrow < $iRow) {
            $highestrow = $iRow;
        }
        $aRows = $row->getCells();
        $amountColmns =  count($aRows);
        $k = 0;
        foreach ( $aRows as $iColumn => $cell){
            if(strtolower($cell->getValue()) == 'set'){
                $setrowfound = 1;
                $icarddeck = 0;
                $iCell++;
                $CardDeck[$setnr]['columnr'] = $iColumn;
                $CardDeck[$setnr]['rownr'] = $iRow;
                $rowbreakers[$iColumn][$iRow] = $setnr;
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
    $iCols = array_keys($rowbreakers);
    for($i = 0 ; $i < count($iCols) ; $i++) {
        $iRows = array_keys($rowbreakers[$iCols[$i]]);
        for($r = 0 ; $r < count($iRows) ; $r++) {
            if(isset($iRows[$r+1])) {
                $set[$rowbreakers[$iCols[$i]][$iRows[$r]]]['maxrownr'] = $iRows[$r+1] - 1;
            } else {
                $set[$rowbreakers[$iCols[$i]][$iRows[$r]]]['maxrownr'] = $highestrow;
            }
            $set[$rowbreakers[$iCols[$i]][$iRows[$r]]]['startrow'] = $iCols[$i];
            if(isset($iCols[$i+1])) {
                $set[$rowbreakers[$iCols[$i]][$iRows[$r]]]['maxendcolumn'] = $iCols[$i+1];
            } else {
                $set[$rowbreakers[$iCols[$i]][$iRows[$r]]]['maxendcolumn'] = NULL;
            }

        }
    }
    $dump = array_keys($CardDeck);
    for($i = 0 ; $i < count($dump) ; $i++) {
        $CardDeck[$dump[$i]]['maxrownr'] = $set[$dump[$i]]['maxrownr'];
        $CardDeck[$dump[$i]]['maxendcolumn'] = $set[$dump[$i]]['maxendcolumn'];
    }

    //now we are going to go through the file again and fill the right cells to the right carddecks
    $amountdecks = count($CardDeck);
    $setcode = '';
    if($amountdecks > 0) {
        foreach ($sheet->getRowIterator() as $iRow => $row) {
            $aRows = $row->getCells();
            foreach ($aRows as $iColumn => $cell) {
                for ($c = 0; $c < $amountdecks; $c++) {
                    if($CardDeck[$c]['columnr'] <= $iColumn  && $CardDeck[$c]['maxendcolumn'] >= $iColumn) {
                        if($CardDeck[$c]['rownr'] <= $iRow && $CardDeck[$c]['maxrownr'] >= $iRow ) {
                            if (!empty($cell->getValue())) {
                                if (strtolower($cell->getValue()) == 'prompt') {
                                    $setcode = 'prompt';
                                }
                                if (strtolower($cell->getValue()) == 'response') {
                                    $setcode = 'response';
                                }
                                switch (true) {
                                    case $iColumn == $CardDeck[$c]['columnr'] + 1 && $iRow == $CardDeck[$c]['rownr'] + 1:
                                        $CardDeck[$c]['carddeckdescription'] = $cell->getValue();
                                        break;
                                    case $iColumn == $CardDeck[$c]['columnr'] + 1:
                                        $CardDeck[$c][$setcode][$iRow]['cardtitle'] = $cell->getValue();
                                        break;
                                    case $iColumn == $CardDeck[$c]['columnr'] + 2:
                                        $CardDeck[$c][$setcode][$iRow]['special'] = $cell->getValue();
                                        break;
                                    case  $iColumn > $CardDeck[$c]['columnr'] + 2 && !empty($cell->getValue() ):
                                        $CardDeck[$c][$setcode][$iRow]['versions'][] = $cell->getValue();
                                        break;
                                }

                            }
                        }
                    }
                }
            }
        }
    }
    print_r($CardDeck);
    $totalCardDeck[$sheet->getName()] = $CardDeck;
}

$reader->close();
