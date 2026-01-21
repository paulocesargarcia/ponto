<?php
require 'vendor/autoload.php';

use League\Csv\Reader;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

$error = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['file'])) {
    $file = $_FILES['file'];

    if ($file['error'] !== UPLOAD_ERR_OK) {
        $error = 'Erro no upload do arquivo.';
    } elseif ($file['size'] > 5 * 1024 * 1024) {
        $error = 'O arquivo deve ter no máximo 5MB.';
    } else {
        $filePath = $file['tmp_name'];

        try {
            $csv = Reader::createFromPath($filePath, 'r');
            $csv->setDelimiter("\t");
            $csv->setHeaderOffset(0);

            $records = $csv->getRecords();
            $collaborators = [];

            foreach ($records as $record) {
                $dept = $record['Departamento'] ?? '';
                $id = $record['ID Colaborador'] ?? '';
                $name = $record['Nome'] ?? '';
                $date = $record['Data'] ?? '';
                $time = $record['Hora'] ?? '';

                if (empty($id) || empty($date)) continue;

                if (!isset($collaborators[$id])) {
                    $collaborators[$id] = [
                        'name' => $name,
                        'days' => []
                    ];
                }

                if (!isset($collaborators[$id]['days'][$date])) {
                    $collaborators[$id]['days'][$date] = [];
                }

                if (count($collaborators[$id]['days'][$date]) < 4) {
                    $collaborators[$id]['days'][$date][] = $time;
                }
            }

            if (empty($collaborators)) {
                $error = 'Arquivo vazio ou em formato inválido.';
            } else {
                generateExcel($collaborators);
                exit;
            }

        } catch (Exception $e) {
            $error = 'Erro ao processar o arquivo: ' . $e->getMessage();
        }
    }
}

function generateExcel($collaborators) {
    $spreadsheet = new Spreadsheet();
    $spreadsheet->removeSheetByIndex(0); // Remove default sheet

    foreach ($collaborators as $id => $data) {
        $name = $data['name'];
        $sheetName = substr(preg_replace('/[\*\?\:\\\\\/\[\]]/', '', $name), 0, 31);
        if (empty($sheetName)) {
            $sheetName = "Colaborador $id";
        }

        // Ensure unique sheet names
        $baseSheetName = $sheetName;
        $counter = 1;
        while ($spreadsheet->sheetNameExists($sheetName)) {
            $sheetName = substr($baseSheetName, 0, 31 - strlen((string)$counter) - 1) . "($counter)";
            $counter++;
        }

        $sheet = $spreadsheet->createSheet();
        $sheet->setTitle($sheetName);

        // Header
        $sheet->setCellValue('A1', 'Nome');
        $sheet->setCellValue('B1', 'Data');
        $sheet->setCellValue('C1', '1');
        $sheet->setCellValue('D1', '2');
        $sheet->setCellValue('E1', '3');
        $sheet->setCellValue('F1', '4');
        $sheet->setCellValue('G1', 'Manhã');
        $sheet->setCellValue('H1', 'Tarde');
        $sheet->setCellValue('I1', 'Total');

        $row = 2;
        foreach ($data['days'] as $date => $times) {
            sort($times); // Ensure times are in order

            $sheet->setCellValue("A$row", $name);

            for ($i = 0; $i < 4; $i++) {
                $col = chr(ord('C') + $i);
                $time = $times[$i] ?? '';
                if ($time) {
                    $parts = explode(':', $time);
                    $excelTime = ($parts[0] * 3600 + ($parts[1] ?? 0) * 60 + ($parts[2] ?? 0)) / 86400;
                    $sheet->setCellValue("$col$row", $excelTime);
                    $sheet->getStyle("$col$row")->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_TIME4);
                } else {
                    $sheet->setCellValue("$col$row", '');
                }
            }

            // Formulas
            // Manhã (G): =IF(AND(Cn<>"",Dn<>""),Dn-Cn,"")
            $sheet->setCellValue("G$row", "=IF(AND(C$row<>\"\",D$row<>\"\"),D$row-C$row,\"\")");

            // Tarde (H): =IF(AND(En<>"",Fn<>""),Fn-En,"")
            $sheet->setCellValue("H$row", "=IF(AND(E$row<>\"\",F$row<>\"\"),F$row-E$row,\"\")");

            // Total (I): =IF(AND(Gn="",Hn=""),"",IF(Gn="",0,Gn)+IF(Hn="",0,Hn))
            $sheet->setCellValue("I$row", "=IF(AND(G$row=\"\",H$row=\"\"),\"\",IF(G$row=\"\",0,G$row)+IF(H$row=\"\",0,H$row))");

            // Formatting
            if ($date) {
                $dParts = explode('/', $date);
                if (count($dParts) === 3) {
                    $excelDate = \PhpOffice\PhpSpreadsheet\Shared\Date::formattedPHPToExcel($dParts[2], $dParts[1], $dParts[0]);
                    $sheet->setCellValue("B$row", $excelDate);
                } else {
                    $sheet->setCellValue("B$row", $date);
                }
            }
            $sheet->getStyle("B$row")->getNumberFormat()->setFormatCode('dd/mm/yyyy');
            $sheet->getStyle("G$row")->getNumberFormat()->setFormatCode('[h]:mm:ss');
            $sheet->getStyle("H$row")->getNumberFormat()->setFormatCode('[h]:mm:ss');
            $sheet->getStyle("I$row")->getNumberFormat()->setFormatCode('[h]:mm:ss');

            $row++;
        }

        // Auto size columns
        foreach (range('A', 'I') as $col) {
            $sheet->getColumnDimension($col)->setAutoSize(true);
        }
    }

    $writer = new Xlsx($spreadsheet);

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="cartao_ponto.xlsx"');
    header('Cache-Control: max-age=0');

    $writer->save('php://output');
}
?>
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Conversor de Cartão-Ponto</title>
    <style>
        body { font-family: sans-serif; margin: 40px; }
        .error { color: red; margin-bottom: 20px; }
        form { border: 1px solid #ccc; padding: 20px; border-radius: 5px; display: inline-block; }
        input[type="file"] { margin-bottom: 10px; }
        button { cursor: pointer; padding: 10px 20px; }
    </style>
</head>
<body>
    <h1>Conversor de Cartão-Ponto (TXT → Excel)</h1>

    <?php if ($error): ?>
        <div class="error"><?php echo htmlspecialchars($error); ?></div>
    <?php endif; ?>

    <form action="index.php" method="post" enctype="multipart/form-data">
        <div>
            <label for="file">Selecione o arquivo TXT/TSV:</label><br><br>
            <input type="file" name="file" id="file" accept=".txt,.tsv" required>
        </div>
        <div>
            <button type="submit">Gerar Excel</button>
        </div>
    </form>
</body>
</html>
