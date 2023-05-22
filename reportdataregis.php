<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Id');
$sheet->setCellValue('B1', 'Jenis Pendaftaran');
$sheet->setCellValue('C1', 'Tanggal Masuk');
$sheet->setCellValue('D1', 'NIS');
$sheet->setCellValue('E1', 'No Peserta Ujian');
$sheet->setCellValue('F1', 'Apakah Pernah Paud');
$sheet->setCellValue('G1', 'Apakah Pernah Tk');
$sheet->setCellValue('H1', 'No SKHUN');
$sheet->setCellValue('I1', 'No Ijazah');
$sheet->setCellValue('J1', 'Hobi');
$sheet->setCellValue('K1', 'Cita - Cita');


$koneksi = mysqli_connect("localhost", "root", "", "formulir");
$query = mysqli_query($koneksi, "SELECT * FROM registrasi");
$i = 2;
$no = 1;
while ($row = mysqli_fetch_array($query)) {
    $sheet->setCellValue('A' . $i, $no++);
    $sheet->setCellValue('B' . $i, $row['jenis_pendaftaran']);
    $sheet->setCellValue('C' . $i, $row['tanggal_masuk']);
    $sheet->setCellValue('D' . $i, $row['nis']);
    $sheet->setCellValue('E' . $i, $row['noPesertaUjian']);
    $sheet->setCellValue('F' . $i, $row['apakah_pernah_paud']);
    $sheet->setCellValue('G' . $i, $row['apakah_pernah_tk']);
    $sheet->setCellValue('H' . $i, $row['noSKHUN']);
    $sheet->setCellValue('I' . $i, $row['noIJAZAH']);
    $sheet->setCellValue('J' . $i, $row['hobi']);
    $sheet->setCellValue('K' . $i, $row['citaCita']);


    $i++;
}

$styleArray = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
];
$sheet->getStyle('A1:K' . ($i - 1))->applyFromArray($styleArray);

$writer = new Xlsx($spreadsheet);
$writer->save('Data Registrasi.xlsx');
?>