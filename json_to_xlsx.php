<?php
// Incluindo o autoload do Composer
require 'vendor/autoload.php';

// Importando as classes necessárias
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Lendo os dados do arquivo JSON
$data = json_decode(file_get_contents('data.json'), true);

$spreadsheet = new Spreadsheet();

// Loop para percorrer os dados
foreach ($data as $item) {
    $sheet = $spreadsheet->createSheet();
    // Usando apenas a data para o título da planilha
    $sheet->setTitle(substr($item['inicio'], 0, 10));

    $headers = array_keys($item);
    $sheet->fromArray([$headers], null, 'A1');

    $row = 2;
    $sheet->fromArray(array_values($item), null, 'A' . $row);
}

$spreadsheet->removeSheetByIndex(0);

// Salvando o arquivo
$filename = 'data.xlsx';

$writer = new Xlsx($spreadsheet);
$temp_file = tempnam(sys_get_temp_dir(), 'xlsx');
$writer->save($temp_file);

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="' . $filename . '"');
header('Content-Length: ' . filesize($temp_file));
readfile($temp_file);

unlink($temp_file);
exit;
?>
