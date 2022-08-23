<?php

use Phppot\DataSource;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

require_once 'DataSource.php';
$db = new DataSource();
$conn = $db->getConnection();
require_once('./vendor/autoload.php');

if (isset($_POST["import"])) {

    $allowedFileType = [
        'application/vnd.ms-excel',
        'text/xls',
        'text/xlsx',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ];

    if (in_array($_FILES["file"]["type"], $allowedFileType)) {

        $targetPath = 'uploads/' . $_FILES['file']['name'];
        move_uploaded_file($_FILES['file']['tmp_name'], $targetPath);

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
}
?>

<!DOCTYPE html>
<html>

<head>
    <style>
        body {
            font-family: Arial;
            width: 550px;
        }

        .outer-container {
            background: #F0F0F0;
            border: #e0dfdf 1px solid;
            padding: 40px 20px;
            border-radius: 2px;
        }

        .btn-submit {
            background: #333;
            border: #1d1d1d 1px solid;
            border-radius: 2px;
            color: #f0f0f0;
            cursor: pointer;
            padding: 5px 20px;
            font-size: 0.9em;
        }

        .tutorial-table {
            margin-top: 40px;
            font-size: 0.8em;
            border-collapse: collapse;
            width: 100%;
        }

        .tutorial-table th {
            background: #f0f0f0;
            border-bottom: 1px solid #dddddd;
            padding: 8px;
            text-align: left;
        }

        .tutorial-table td {
            background: #FFF;
            border-bottom: 1px solid #dddddd;
            padding: 8px;
            text-align: left;
        }

        #response {
            padding: 10px;
            margin-top: 10px;
            border-radius: 2px;
            display: none;
        }

        .success {
            background: #c7efd9;
            border: #bbe2cd 1px solid;
        }

        .error {
            background: #fbcfcf;
            border: #f3c6c7 1px solid;
        }

        div#response.display-block {
            display: block;
        }
    </style>
</head>

<body>
    <h2>Tool Import Excel File into MySQL Database</h2>

    <div class="outer-container">
        <form action="" method="post" name="frmExcelImport" id="frmExcelImport" enctype="multipart/form-data">
            <div>
                <label>Choose Excel File</label> <input type="file" name="file" id="file" accept=".xls,.xlsx">
                <button type="submit" id="submit" name="import" class="btn-submit">Import</button>
            </div>
        </form>
    </div>
    <div id="response" class="<?php if (!empty($type)) {
                                    echo $type . " display-block";
                                } ?>">
        <?php
        if (!empty($message)) {
            echo $message . PHP_EOL;
        }

        echo "<br />Total codes insert: ".count($successCode) . PHP_EOL;
        var_dump($successCode);
        ?>
    </div>


    <?php
    $sqlSelect = "SELECT * FROM promotions";
    $result = $db->select($sqlSelect);
    if (!!empty($result))
    {
    ?>

        <table class='tutorial-table'>
            <thead>
                <tr>
                    <th>ID</th>
                    <th>Code</th>
                    <th>Name</th>

                </tr>
            </thead>
            <?php
            foreach ($result as $row) { // ($row = mysqli_fetch_array($result))
            ?>
                <tbody>
                    <tr>
                        <td><?php echo $row['PromotionId']; ?></td>
                        <td><?php echo $row['PromotionCode']; ?></td>
                        <td><?php echo $row['PromotionName']; ?></td>
                    </tr>
                <?php
            }
                ?>
                </tbody>
        </table>
    <?php
    }
    ?>

</body>

</html>