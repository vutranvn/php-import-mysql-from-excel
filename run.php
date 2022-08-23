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

    $successCode = [];
    for ($i = 0; $i <= $sheetCount; $i++) {
        $code = "";
        if (isset($spreadSheetAry[$i][0])) {
            $code = mysqli_real_escape_string($conn, $spreadSheetAry[$i][0]);
        }

        if (!empty($code)) {
            $query = "insert into promotions(PromotionCode,PromotionName,PromotionNameEng,PromotionContent,PromotionContentEng,BeginDate,EndDate,MinCost,LimitAll,LimitPerUser,ApplyFrom,IsAllCustomer,IsAllService,StatusId,OrderPercent,OrderPaidVN,PromotionImage,CustomerIds,CrUserId,CrDateTime) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
            $paramType = "sssssssdddddddddssds";

            $paramArray = array(
                $code,
                'GRABWU',
                'GRABWU',
                '<p></p>',
                '<p></p>',
                '2022-08-22',
                null,
                0,
                1,
                0,
                3,
                1,
                0,
                2,
                100,
                0,
                '',
                null,
                1,
                date('Y-m-d H:i:s')
            );

            $insertId = $db->insert($query, $paramType, $paramArray);
            // $query = "insert into tbl_test(name,description) values('" . $name . "','" . $description . "')";
            $result1 = mysqli_query($conn, $query);

            if (!empty($insertId)) {
                $successCode[] = $insertId . '-' .$code;
                // Update promotion for products (services)
                $listProducts = [111, 116, 121, 128, 109];
                foreach( $listProducts as $productId) {
                    $query = "insert into promotionproducts(PromotionId,ProductId) values('" . $insertId . "','" . $productId . "')";
                    $result = mysqli_query($conn, $query);
                }

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

print_r($successCode);
echo "Total codes insert: ".count($successCode) . PHP_EOL;
?>