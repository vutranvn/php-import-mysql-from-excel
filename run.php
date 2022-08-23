<?php

use Phppot\DataSource;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

require_once 'DataSource.php';
$db = new DataSource();
$conn = $db->getConnection();
require_once('./vendor/autoload.php');


$targetPath = 'uploads/codes.xlsx';

if (file_exists($targetPath)) {
    $Reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();

    $spreadSheet = $Reader->load($targetPath);
    $excelSheet = $spreadSheet->getActiveSheet();
    $spreadSheetAry = $excelSheet->toArray();
    $sheetCount = count($spreadSheetAry);

    $successItems = [];
    for ($i = 0; $i <= $sheetCount; $i++) {
        $name = '';
        $description = '';
        if (isset($spreadSheetAry[$i][0])) {
            $name = mysqli_real_escape_string($conn, $spreadSheetAry[$i][0]);
        }

        if (isset($spreadSheetAry[$i][0])) {
            $description = mysqli_real_escape_string($conn, $spreadSheetAry[$i][0]);
        }

        if (!empty($name) || !empty($description)) {
            $query = "insert into tbl_test(name,descriptipn) values(?,?)";
            $paramType = "ss";

            $paramArray = array(
                $code,
                $description
            );

            $insertId = $db->insert($query, $paramType, $paramArray);
            // $query = "insert into tbl_test(name,description) values('" . $name . "','" . $description . "')";
            $result = mysqli_query($conn, $query);

            if (!empty($insertId)) {
                $successItems[] = $insertId . '-' .$code;

                $type = "success";
                $message = "Excel Data Imported into the Database";
            } else {
                $type = "error";
                $message = "Problem in Importing Excel Data";
            }
        }
    }

} else {
    $type = "error";
    $message = "Invalid File Type. Upload Excel File.";
}

if (!empty($message)) {
    echo $message . PHP_EOL;
}

print_r($successItems);
echo "Total items insert: ".count($successItems) . PHP_EOL;
?>