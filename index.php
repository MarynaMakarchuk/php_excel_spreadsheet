<?php
/*Создать в базе таблицу товаров, с нужными полями на ваше усмотрение и оформить вывод прайс-листа в ексель.
Отформатировать своими стилями, придать прайсу законченный вид.
 */
require '../vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$user = 'root';
$pass = '';
try {
    $dbh = new PDO('mysql:host=localhost;dbname=classicmodels', $user, $pass);
    $result = $dbh->query('SELECT productCode, productName, buyPrice from products');

   $i = 0;
    while ($row = $result->fetch()) {

        $productCode[$i] = $row['productCode'];
        $productName[$i] = $row['productName'];
        $buyPrice[$i] = $row['buyPrice'];
        $i++;
    }
    $dbh = null;
} catch (PDOException $e) {
    print "Error!: " . $e->getMessage() . "<br/>";
    die();
}


function createPriceList($productCode, $productName, $buyPrice)
{

    $i = 1;
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $i = 1;
    foreach ($productCode as $code) {
        $i++;
        $sheet->setCellValue('A1', 'ProductCode');
        $sheet->getStyle('A1')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('A1')->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
        $sheet->getStyle('A1')->getFont()->setBold(true);
        $sheet->setCellValue('A' . $i, $code);
        $sheet->getStyle('A' . $i)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
        $styleArray = [
            'borders' => [
                'outline' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
            ],
        ];
        $sheet->getStyle('A1')->applyFromArray($styleArray);
        $sheet->getStyle('A' . $i)->applyFromArray($styleArray);
    }
    $i = 1;
    foreach ($productName as $name) {
        $i++;
        $sheet->setCellValue('B1', 'ProductName');
        $sheet->getStyle('B1')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('B1')->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
        $sheet->getStyle('B1')->getFont()->setBold(true);
        $sheet->setCellValue('B' . $i, $name);
        $sheet->getStyle('B' . $i)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
        $styleArray = [
            'borders' => [
                'outline' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
            ],
        ];
        $sheet->getStyle('B1')->applyFromArray($styleArray);
        $sheet->getStyle('B' . $i)->applyFromArray($styleArray);
    }
    $i = 1;
    foreach ($buyPrice as $price) {
        $i++;
        $sheet->setCellValue('C1', 'Price');
        $sheet->getStyle('C1')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('C1')->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
        $sheet->getStyle('C1')->getFont()->setBold(true);
        $sheet->setCellValue('C' . $i, $price);
        $sheet->getStyle('C' . $i)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
        $styleArray = [
            'borders' => [
                'outline' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
            ],
        ];
        $sheet->getStyle('C1')->applyFromArray($styleArray);
        $sheet->getStyle('C' . $i)->applyFromArray($styleArray);
    }
    $writer = new Xlsx($spreadsheet);
    $writer->save('price.xlsx');
}
createPriceList($productCode, $productName, $buyPrice);






