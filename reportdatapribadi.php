<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Id');
$sheet->setCellValue('B1', 'Nama');
$sheet->setCellValue('C1', 'Jenis Kelamin');
$sheet->setCellValue('D1', 'NISN');
$sheet->setCellValue('E1', 'NIK');
$sheet->setCellValue('F1', 'Tempat Lahir');
$sheet->setCellValue('G1', 'Tanggal Lahir');
$sheet->setCellValue('H1', 'Agama');
$sheet->setCellValue('I1', 'Berkebutuhan Khusus');
$sheet->setCellValue('J1', 'Alamat Jalan');
$sheet->setCellValue('K1', 'RT');
$sheet->setCellValue('L1', 'RW');
$sheet->setCellValue('M1', 'Nama Dusun');
$sheet->setCellValue('N1', 'Nama Kelurahan Desa');
$sheet->setCellValue('O1', 'Kecamatan');
$sheet->setCellValue('P1', 'Kode Pos');
$sheet->setCellValue('Q1', 'Tempat Tinggal');
$sheet->setCellValue('R1', 'Moda Transportasi');
$sheet->setCellValue('S1', 'Nomor HP');
$sheet->setCellValue('T1', 'Email');
$sheet->setCellValue('U1', 'Penerima KIP');
$sheet->setCellValue('V1', 'No KIP');
$sheet->setCellValue('W1', 'Kewarganegaraan');
$sheet->setCellValue('X1', 'Negara');



$koneksi = mysqli_connect("localhost", "root", "", "formulir");
$query = mysqli_query($koneksi, "SELECT * FROM data_pribadi");
$i = 2;
$no = 1;
while ($row = mysqli_fetch_array($query)) {
    $sheet->setCellValue('A' . $i, $no++);
    $sheet->setCellValue('B' . $i, $row['nama']);
    $sheet->setCellValue('C' . $i, $row['jenis_kelamin']);
    $sheet->setCellValue('D' . $i, $row['nisn']);
    $sheet->setCellValue('E' . $i, $row['nik']);
    $sheet->setCellValue('F' . $i, $row['tempat_lahir']);
    $sheet->setCellValue('G' . $i, $row['tanggal_lahir']);
    $sheet->setCellValue('H' . $i, $row['agama']);
    $sheet->setCellValue('I' . $i, $row['berkebutuhan_khusus']);
    $sheet->setCellValue('J' . $i, $row['alamat_jalan']);
    $sheet->setCellValue('K' . $i, $row['rt']);
    $sheet->setCellValue('L' . $i, $row['rw']);
    $sheet->setCellValue('M' . $i, $row['nama_dusun']);
    $sheet->setCellValue('N' . $i, $row['nama_kelurahan_desa']);
    $sheet->setCellValue('O' . $i, $row['kecamatan']);
    $sheet->setCellValue('P' . $i, $row['kode_pos']);
    $sheet->setCellValue('Q' . $i, $row['tempat_tinggal']);
    $sheet->setCellValue('R' . $i, $row['moda_transportasi']);
    $sheet->setCellValue('S' . $i, $row['nomor_hp']);
    $sheet->setCellValue('T' . $i, $row['email']);
    $sheet->setCellValue('U' . $i, $row['penerima_kip']);
    $sheet->setCellValue('V' . $i, $row['no_kip']);
    $sheet->setCellValue('W' . $i, $row['kewarganegaraan']);
    $sheet->setCellValue('X' . $i, $row['negara']);


    $i++;
}

$styleArray = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
];
$sheet->getStyle('A1:X' . ($i - 1))->applyFromArray($styleArray);

$writer = new Xlsx($spreadsheet);
$writer->save('Data Pribadi.xlsx');
?>