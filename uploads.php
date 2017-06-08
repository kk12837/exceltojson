<?php
/**
 * Created by PhpStorm.
 * User: josiesuo
 * Date: 8/5/2017
 * Time: 14:20
 */

//remember to change the dir of your working environment
$target_dir = 'C:/xampp/htdocs/keith/exceltojson/uploads/';
$download_dir = 'C:/xampp/htdocs/keith/exceltojson/downloads/';
$jsonFilename = 'C:/xampp/htdocs/keith/exceltojson/downloads/t_words.json';
$jsonGroupFilename = 'C:/xampp/htdocs/keith/exceltojson/downloads/t_groups.json';

$arrAllFileName = array();
$arrTargetFiles = array();
$arrFileBaseName = array();
if ($_SERVER['REQUEST_METHOD'] == 'POST') {
    foreach ($_FILES['filesToUpload']['name'] as $i => $name) {
        if (strlen($_FILES['filesToUpload']['name'][$i]) > 1) {

            if (!file_exists($target_dir)) {
                mkdir($target_dir, 0777, true);
            }

            $filename = $_FILES["filesToUpload"]["name"][$i];
            $fileBaseName = basename($_FILES["filesToUpload"]["name"][$i], '.xls');
            $target_file = $target_dir . basename($_FILES["filesToUpload"]["name"][$i]);

            if (move_uploaded_file($_FILES['filesToUpload']['tmp_name'][$i], $target_file)) {
                $arrAllFileName[] = $filename;
                $arrTargetFiles[] = $target_file;
                $arrFileBaseName[] = $fileBaseName;

            }
        }
    }
}

//get library
$arrayMain = array();
$arrayGroup = array();
$id = 1;
require_once 'PHPExcel-1.8/Classes/PHPExcel.php';
require_once 'PHPExcel-1.8/Classes/PHPExcel/IOFactory.php';

if(isset($_POST["submit"]) && count($arrAllFileName) == count($arrTargetFiles) ) {

    foreach ($arrAllFileName as $key => $filename) {


        $fileType = pathinfo($arrTargetFiles[$key],PATHINFO_EXTENSION);
        // Allow certain file formats
        if($fileType != "xls") {
            echo "Sorry, only xls files is allowed.";
            return;
        }
        echo "<br/>";
        echo "You uploaded ".$filename."\n";

        //set free = 1 if it is suyu
        $free = 0;
        if ($filename == 'SuYu.xls') {
            $free = 1;
        } else {
            $free = 0;
        }

        $inputFileType = 'Excel5';
        $inputFileName = $arrTargetFiles[$key];

        $objReader = PHPExcel_IOFactory::createReader($inputFileType);
        $objPHPExcel = $objReader->load($inputFileName);

        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'HTML');
        $objWriter->save('php://output');


        $filename = $arrTargetFiles[$key];
        $type = PHPExcel_IOFactory::identify($filename);
        $objReader = PHPExcel_IOFactory::createReader($type);
        $objPHPExcel = $objReader->load($filename);

        $i = 1;
        $arrayOutput = array();
        $arrayGroupOutput = array();

        foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {
            $worksheets[$worksheet->getTitle()] = $worksheet->toArray();
            foreach ($worksheet->toArray() as $data) {

                if (is_string($data[0])) {
                    $topic = explode("(", $data[0]);

                    if (count($topic) > 1) {
                        $topic = explode(")", $topic[1]);
                        $topic = $topic[0];

                        $arrayGroupOutput['id'] = $key+1;
                        $arrayGroupOutput['name_hans'] = $topic;
                        $arrayGroupOutput['name_hant'] = $topic;
                        $arrayGroupOutput['icon'] = 'null';//'icon'.$key.'.png';
                        $arrayGroupOutput['free'] = $free;
                        $arrayGroup['RECORDS'][] = $arrayGroupOutput;
                    }

                }

                if (count($data) < 3) {
                    continue;
                }

                //set free = 1 if it is suyu
                $suyu = 0;
                if ($arrFileBaseName[$key] == 'SuYu') {
                    $suyu = 1;
                }

                if (!empty($data) && $i > 3) {
                    //var_dump($data);
                    if ($data[1] && $data[2] && $data[3] && $data[4] && $data[5]) {
                        $arrayOutput['id'] = $id;
                        $arrayOutput['text_can_hans'] = $data[5];
                        $arrayOutput['text_can_hant'] = $data[2];
                        $arrayOutput['text_man_hans'] = $data[4];
                        $arrayOutput['text_man_hant'] = $data[1];
                        $arrayOutput['phonetic_can'] = checkString($data[3]);
                        $arrayOutput['record_file'] = 'file'.$id;
                        $arrayOutput['group_id'] = $key+1;
                        $arrayOutput['free'] = $suyu;
                        $arrayOutput['study'] = 0;
                        $id++;
                        $arrayMain['RECORDS'][] = $arrayOutput;
                    }

                    if (count($data) < 7) {
                        continue;
                    }

                    if ($data[7] && $data[8] && $data[9] && $data[10] && $data[11]) {
                        $arrayOutput['id'] = $id;
                        $arrayOutput['text_can_hans'] = $data[11];
                        $arrayOutput['text_can_hant'] = $data[8];
                        $arrayOutput['text_man_hans'] = $data[10];
                        $arrayOutput['text_man_hant'] = $data[7];
                        $arrayOutput['phonetic_can'] = checkString($data[9]);
                        $arrayOutput['record_file'] = 'file'.$id;
                        $arrayOutput['group_id'] = $key+1;
                        $arrayOutput['free'] = $suyu;
                        $arrayOutput['study'] = 0;
                        $id++;
                        $arrayMain['RECORDS'][] = $arrayOutput;
                    }

                    //break;
                }

                $i++;
            }


        }


    }

}

//save data to json
if (!empty($arrayMain) && !empty($arrayGroup)) {

    if (!file_exists($download_dir)) {
        mkdir($download_dir, 0777, true);
    }

    $fp = fopen($jsonFilename, 'w');
    fwrite($fp, json_encode($arrayMain, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES ));
    fclose($fp);

    echo "<p>Download t_words.json file</p>
            <a target='_blank\' href='http://127.0.0.1/keith/exceltojson/downloads/t_words.json'>Download t_words.json Here!</a>";
    echo "<br/>";

    $fp = fopen($jsonGroupFilename, 'w');
    fwrite($fp, json_encode($arrayGroup, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES ));
    fclose($fp);

    echo "<p>Download t_groups.json file</p>
            <a target='_blank\' href='http://127.0.0.1/keith/exceltojson/downloads/t_groups.json'>Download t_groups.json Here!</a>";
    echo "<br/>";

}

echo "<br/>";
echo "<a href='http://127.0.0.1/keith/exceltojson/'>Return to homepage</a>";

function checkString($string) {
    $returnString = $string;
    if ($string) {
        $returnString = str_replace(' ', '&', $string);
    }
    return $returnString;
}
