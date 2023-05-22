<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Id');
$sheet->setCellValue('B1', 'Nama');
$sheet->setCellValue('C1', 'Tahun Lahir');
$sheet->setCellValue('D1', 'Pendidikan');
$sheet->setCellValue('E1', 'Pekerjaan');
$sheet->setCellValue('F1', 'Penghasilan Bulan');
$sheet->setCellValue('G1', 'Berkebutuhan Khusus');


$koneksi = mysqli_connect("localhost", "root", "", "formulir");
$query = mysqli_query($koneksi, "SELECT * FROM ibu_kandung");
$i = 2;
$no = 1;
while ($row = mysqli_fetch_array($query)) {
    $sheet->setCellValue('A' . $i, $no++);
    $sheet->setCellValue('B' . $i, $row['nama']);
    $sheet->setCellValue('C' . $i, $row['tahun_lahir']);
    $sheet->setCellValue('D' . $i, $row['pendidikan']);
    $sheet->setCellValue('E' . $i, $row['pekerjaan']);
    $sheet->setCellValue('F' . $i, $row['penghasilan_bulan']);
    $sheet->setCellValue('G' . $i, $row['berkebutuhan_khusus']);

    $i++;
}

$styleArray = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
];
$sheet->getStyle('A1:G' . ($i - 1))->applyFromArray($styleArray);

$writer = new Xlsx($spreadsheet);
$writer->save('Data Ibu.xlsx');
?>