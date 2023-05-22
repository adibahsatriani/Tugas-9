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


$sheet->setCellValue('A7', 'Id');
$sheet->setCellValue('B7', 'Nama');
$sheet->setCellValue('C7', 'Jenis Kelamin');
$sheet->setCellValue('D7', 'NISN');
$sheet->setCellValue('E7', 'NIK');
$sheet->setCellValue('F7', 'Tempat Lahir');
$sheet->setCellValue('G7', 'Tanggal Lahir');
$sheet->setCellValue('H7', 'Agama');
$sheet->setCellValue('I7', 'Berkebutuhan Khusus');
$sheet->setCellValue('J7', 'Alamat Jalan');
$sheet->setCellValue('K7', 'RT');
$sheet->setCellValue('L7', 'RW');
$sheet->setCellValue('M7', 'Nama Dusun');
$sheet->setCellValue('N7', 'Nama Kelurahan Desa');
$sheet->setCellValue('O7', 'Kecamatan');
$sheet->setCellValue('P7', 'Kode Pos');
$sheet->setCellValue('Q7', 'Tempat Tinggal');
$sheet->setCellValue('R7', 'Moda Transportasi');
$sheet->setCellValue('S7', 'Nomor HP');
$sheet->setCellValue('T7', 'Email');
$sheet->setCellValue('U7', 'Penerima KIP');
$sheet->setCellValue('V7', 'No KIP');
$sheet->setCellValue('W7', 'Kewarganegaraan');
$sheet->setCellValue('X7', 'Negara');

$sheet->setCellValue('A13', 'Id');
$sheet->setCellValue('B13', 'Nama');
$sheet->setCellValue('C13', 'Tahun Lahir');
$sheet->setCellValue('D13', 'Pendidikan');
$sheet->setCellValue('E13', 'Pekerjaan');
$sheet->setCellValue('F13', 'Penghasilan Bulan');
$sheet->setCellValue('G13', 'Berkebutuhan Khusus');


$sheet->setCellValue('A19', 'Id');
$sheet->setCellValue('B19', 'Nama');
$sheet->setCellValue('C19', 'Tahun Lahir');
$sheet->setCellValue('D19', 'Pendidikan');
$sheet->setCellValue('E19', 'Pekerjaan');
$sheet->setCellValue('F19', 'Penghasilan Bulan');
$sheet->setCellValue('G19', 'Berkebutuhan Khusus');



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

$koneksi = mysqli_connect("localhost", "root", "", "formulir");
$query = mysqli_query($koneksi, "SELECT registrasi.*, data_pribadi.*
                                FROM registrasi
                                JOIN data_pribadi ON registrasi.id_regis = data_pribadi.id_pribadi");
$i = 8;
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
$sheet->getStyle('A7:X' . ($i - 1))->applyFromArray($styleArray);


$koneksi = mysqli_connect("localhost", "root", "", "formulir");
$query = mysqli_query($koneksi, "SELECT registrasi.*, data_pribadi.*, ayah_kandung.*
                                FROM registrasi
                                JOIN data_pribadi ON registrasi.id_regis = data_pribadi.id_pribadi
                                JOIN ayah_kandung ON registrasi.id_regis = ayah_kandung.id_ayah");
$i = 14;
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
$sheet->getStyle('A13:G' . ($i - 1))->applyFromArray($styleArray);


$koneksi = mysqli_connect("localhost", "root", "", "formulir");
$query = mysqli_query($koneksi, "SELECT registrasi.*, data_pribadi.*, ayah_kandung.*, ibu_kandung.*
                                FROM registrasi
                                JOIN data_pribadi ON registrasi.id_regis = data_pribadi.id_pribadi
                                JOIN ayah_kandung ON registrasi.id_regis = ayah_kandung.id_ayah
                                JOIN ibu_kandung ON registrasi.id_regis = ibu_kandung.id_ibu");
$i = 20;
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
$sheet->getStyle('A19:G' . ($i - 1))->applyFromArray($styleArray);

$writer = new Xlsx($spreadsheet);
$writer->save('reportall.xlsx');
?>