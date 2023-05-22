<?php
include 'koneksi.php';
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'No');
$sheet->setCellValue('B1', 'Jenis Pendaftaran');
$sheet->setCellValue('C1', 'Tgl Masuk Sekolah');
$sheet->setCellValue('D1', 'NIS');
$sheet->setCellValue('E1', 'No. Peserta Ujian');
$sheet->setCellValue('F1', 'Pernah Paud');
$sheet->setCellValue('G1', 'Pernah Tk');
$sheet->setCellValue('H1', 'No.Skhun');
$sheet->setCellValue('I1', 'No.Ijazah');
$sheet->setCellValue('J1', 'Hobi');
$sheet->setCellValue('K1', 'Cita-Cita');
$sheet->setCellValue('L1', 'Nama Lengkap');
$sheet->setCellValue('M1', 'Jk');
$sheet->setCellValue('N1', 'NISN');
$sheet->setCellValue('O1', 'NIK');
$sheet->setCellValue('P1', 'Tempat Lahir');
$sheet->setCellValue('Q1', 'Tanggal Lahir');
$sheet->setCellValue('R1', 'Agama');
$sheet->setCellValue('S1', 'Berkebutuhan Khusus');
$sheet->setCellValue('T1', 'Alamat');
$sheet->setCellValue('U1', 'RT');
$sheet->setCellValue('V1', 'RW');
$sheet->setCellValue('W1', 'Dusun');
$sheet->setCellValue('X1', 'Kelurahan');
$sheet->setCellValue('Y1', 'Kecamatan');
$sheet->setCellValue('Z1', 'Kode Pos');
$sheet->setCellValue('AA1', 'Tempat Tinggal');
$sheet->setCellValue('AB1', 'Transportasi');
$sheet->setCellValue('AC1', 'No.Hp');
$sheet->setCellValue('AD1', 'No.Tlp');
$sheet->setCellValue('AE1', 'Email');
$sheet->setCellValue('AF1', 'Penerima Kps');
$sheet->setCellValue('AG1', 'No.Kps');
$sheet->setCellValue('AH1', 'Kewarganegaraan');
$sheet->setCellValue('AI1', 'Nama Ayah Kandung');
$sheet->setCellValue('AJ1', 'Tahun Lahir');
$sheet->setCellValue('AK1', 'Pendidikan');
$sheet->setCellValue('AL1', 'Pekerjaan');
$sheet->setCellValue('AM1', 'Penghasilan Bulanan');
$sheet->setCellValue('AN1', 'Berkebutuhan Khusus');
$sheet->setCellValue('AO1', 'Nama Ibu Kandung');
$sheet->setCellValue('AP1', 'Tahun Lahir');
$sheet->setCellValue('AQ1', 'Pendidikan');
$sheet->setCellValue('AR1', 'Pekerjaan');
$sheet->setCellValue('AS1', 'Penghasilan');
$sheet->setCellValue('AT1', 'Berkebutuhan Khusus');

$query = mysqli_query($koneksi, "SELECT * FROM regis, datapribadi, d_ayah, dibu");
$i = 2;
$no = 1;
while ($row = mysqli_fetch_array($query)) {
    $sheet->setCellValue('A' . $i, $no++);
    $sheet->setCellValue('B' . $i, $row['jenis_pendaftaran']);
    $sheet->setCellValue('C' . $i, $row['tgl_msksklh']);
    $sheet->setCellValue('D' . $i, $row['NIS']);
    $sheet->setCellValue('E' . $i, $row['nomor_pesertaujn']);
    $sheet->setCellValue('F' . $i, $row['paud']);
    $sheet->setCellValue('G' . $i, $row['tk']);
    $sheet->setCellValue('H' . $i, $row['skhun']);
    $sheet->setCellValue('I' . $i, $row['ijazah']);
    $sheet->setCellValue('J' . $i, $row['hobi']);
    $sheet->setCellValue('K' . $i, $row['cita']);
    $sheet->setCellValue('L' . $i, $row['nama']);
    $sheet->setCellValue('M' . $i, $row['jenis_kelamin']);
    $sheet->setCellValue('N' . $i, $row['nisn']);
    $sheet->setCellValue('O' . $i, $row['nik']);
    $sheet->setCellValue('P' . $i, $row['tempat_lahir']);
    $sheet->setCellValue('Q' . $i, $row['tgl_lahir']);
    $sheet->setCellValue('R' . $i, $row['agama']);
    $sheet->setCellValue('S' . $i, $row['abk']);
    $sheet->setCellValue('T' . $i, $row['alamat']);
    $sheet->setCellValue('U' . $i, $row['rt']);
    $sheet->setCellValue('V' . $i, $row['rw']);
    $sheet->setCellValue('W' . $i, $row['dusun']);
    $sheet->setCellValue('X' . $i, $row['desa']);
    $sheet->setCellValue('Y' . $i, $row['kecamatan']);
    $sheet->setCellValue('Z' . $i, $row['kodepos']);
    $sheet->setCellValue('AA' . $i, $row['tempat_tinggal']);
    $sheet->setCellValue('AB' . $i, $row['transport']);
    $sheet->setCellValue('AC' . $i, $row['nohp']);
    $sheet->setCellValue('AD' . $i, $row['notelp']);
    $sheet->setCellValue('AE' . $i, $row['email']);
    $sheet->setCellValue('AF' . $i, $row['bantuan_kip']);
    $sheet->setCellValue('AG' . $i, $row['nokip']);
    $sheet->setCellValue('AH' . $i, $row['kewarganegaraan']);
    $sheet->setCellValue('AI' . $i, $row['nama_ayah']);
    $sheet->setCellValue('AJ' . $i, $row['tahun_lahir']);
    $sheet->setCellValue('AK' . $i, $row['pendidikan']);
    $sheet->setCellValue('AL' . $i, $row['pekerjaan']);
    $sheet->setCellValue('AM' . $i, $row['penghasilan']);
    $sheet->setCellValue('AN' . $i, $row['ayah_abk']);
    $sheet->setCellValue('AO' . $i, $row['nama_ibu']);
    $sheet->setCellValue('AP' . $i, $row['tahun_lahir']);
    $sheet->setCellValue('AQ' . $i, $row['pendidikan']);
    $sheet->setCellValue('AR' . $i, $row['pekerjaan']);
    $sheet->setCellValue('AS' . $i, $row['penghasilan']);
    $sheet->setCellValue('AT' . $i, $row['ibu_abk']);
    $i++;
}
$styleArray = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
];

$i = $i - 1;
$sheet->getStyle('A1:AT1' . $i)->applyFromArray($styleArray);
$writer = new Xlsx($spreadsheet);
$writer->save('Report Data Pendaftaran Peserta Didik.xlsx');
?>
<script>
    alert("Mengekspor Data Peserta Didik Ke Excel");
</script>