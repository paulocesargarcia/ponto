<?php
ob_start();
require 'vendor/autoload.php';

use League\Csv\Reader;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

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

            $headers = $csv->getHeader();
            $records = $csv->getRecords();
            $collaborators = [];

            // Mapping for headers
            $map = [
                'dept' => ['Departamento', 'Dept.'],
                'id' => ['ID Colaborador', 'User ID', 'Enroll ID'],
                'name' => ['Nome', 'Name'],
                'date' => ['Data', 'Date'],
                'time' => ['Hora', 'Time'],
                'm1' => ['1'],
                'm2' => ['2'],
                'm3' => ['3'],
                'm4' => ['4'],
            ];

            $findHeader = function($keys) use ($headers) {
                foreach ($keys as $key) {
                    $index = array_search($key, $headers);
                    if ($index !== false) return $key;
                }
                return null;
            };

            $hId = $findHeader($map['id']);
            $hName = $findHeader($map['name']);
            $hDate = $findHeader($map['date']);
            $hTime = $findHeader($map['time']);
            $hM1 = $findHeader($map['m1']);
            $hM2 = $findHeader($map['m2']);
            $hM3 = $findHeader($map['m3']);
            $hM4 = $findHeader($map['m4']);

            foreach ($records as $record) {
                $id = $record[$hId] ?? '';
                $name = $record[$hName] ?? '';
                $date = $record[$hDate] ?? '';

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

                // If explicit marking columns exist
                if ($hM1 || $hM2 || $hM3 || $hM4) {
                    $marks = [];
                    if (!empty($record[$hM1])) $marks[] = $record[$hM1];
                    if (!empty($record[$hM2])) $marks[] = $record[$hM2];
                    if (!empty($record[$hM3])) $marks[] = $record[$hM3];
                    if (!empty($record[$hM4])) $marks[] = $record[$hM4];
                    $collaborators[$id]['days'][$date] = array_slice($marks, 0, 4);
                } else {
                    // Fallback to one marking per row
                    $time = $record[$hTime] ?? '';
                    if (!empty($time) && count($collaborators[$id]['days'][$date]) < 4) {
                        $collaborators[$id]['days'][$date][] = $time;
                    }
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
        $sheetName = mb_substr(preg_replace('/[\*\?\:\\\\\/\[\]]/', '', $name), 0, 31);
        if (empty($sheetName)) {
            $sheetName = "Colaborador $id";
        }

        // Ensure unique sheet names
        $baseSheetName = $sheetName;
        $counter = 1;
        while ($spreadsheet->sheetNameExists($sheetName)) {
            $sheetName = mb_substr($baseSheetName, 0, 31 - strlen((string)$counter) - 1) . "($counter)";
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

        $sheet->getStyle('A1:I1')->getFont()->setBold(true);

        $row = 2;
        $startRow = 2;
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
                    $sheet->getStyle("$col$row")->getNumberFormat()->setFormatCode('h:mm:ss');
                } else {
                    $sheet->setCellValue("$col$row", '');
                }
            }

            // Formulas
            $sheet->setCellValue("G$row", "=IF(AND(C$row<>\"\",D$row<>\"\"),D$row-C$row,\"\")");
            $sheet->setCellValue("H$row", "=IF(AND(E$row<>\"\",F$row<>\"\"),F$row-E$row,\"\")");
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

        $endDataRow = $row - 1;

        // Add Total row
        $row++; // Empty row
        $sheet->setCellValue("I$row", "=SUM(I$startRow:I$endDataRow)");
        $sheet->getStyle("I$row")->getNumberFormat()->setFormatCode('[h]:mm:ss');
        $sheet->getStyle("I$row")->getFont()->setBold(true);

        // Auto size columns
        foreach (range('A', 'I') as $col) {
            $sheet->getColumnDimension($col)->setAutoSize(true);
        }
    }

    $writer = new Xlsx($spreadsheet);

    // Create temporary file to get content length
    $tempFile = tempnam(sys_get_temp_dir(), 'xlsx');
    $writer->save($tempFile);
    $content = file_get_contents($tempFile);
    $size = filesize($tempFile);
    unlink($tempFile);

    if (ob_get_length()) ob_end_clean();

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="cartao_ponto.xlsx"');
    header('Cache-Control: max-age=0');
    header('Expires: Mon, 26 Jul 1997 05:00:00 GMT');
    header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');
    header('Cache-Control: cache, must-revalidate');
    header('Pragma: public');
    header('Content-Length: ' . $size);

    echo $content;
}
?>
<!DOCTYPE html>
<html lang="pt-br" class="h-full bg-slate-50">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Conversor de Cartão-Ponto</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Inter', sans-serif; }
    </style>
</head>
<body class="h-full flex flex-col items-center justify-center p-6">
    <div class="w-full max-w-md">
        <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="p-8">
                <div class="flex flex-col items-center mb-8">
                    <div class="w-12 h-12 bg-indigo-600 rounded-lg flex items-center justify-center mb-4">
                        <svg class="w-6 h-6 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
                        </svg>
                    </div>
                    <h1 class="text-2xl font-bold text-slate-900 tracking-tight text-center">Conversor de Cartão-Ponto</h1>
                    <p class="text-slate-500 mt-2 text-center text-sm">Transforme seus arquivos TXT em planilhas Excel formatadas.</p>
                </div>

                <?php if ($error): ?>
                    <div class="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg flex items-start gap-3">
                        <svg class="w-5 h-5 text-red-500 mt-0.5 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path>
                        </svg>
                        <p class="text-sm text-red-700 font-medium"><?php echo htmlspecialchars($error); ?></p>
                    </div>
                <?php endif; ?>

                <form action="index.php" method="post" enctype="multipart/form-data" class="space-y-6">
                    <div>
                        <label for="file" class="block text-sm font-semibold text-slate-700 mb-2">Selecione o arquivo de ponto</label>
                        <div class="relative group">
                            <input type="file" name="file" id="file" required
                                class="block w-full text-sm text-slate-500
                                file:mr-4 file:py-2.5 file:px-4
                                file:rounded-lg file:border-0
                                file:text-sm file:font-semibold
                                file:bg-indigo-50 file:text-indigo-700
                                hover:file:bg-indigo-100
                                border border-slate-200 rounded-lg p-1.5
                                group-hover:border-indigo-300 transition-colors
                                focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500">
                        </div>
                        <p class="mt-2 text-xs text-slate-400 italic">Formatos suportados: TXT, TSV (separado por TAB).</p>
                    </div>

                    <button type="submit"
                        class="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-semibold py-3 px-4 rounded-lg shadow-sm transition-all duration-200 flex items-center justify-center gap-2 group focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:ring-offset-2">
                        <span>Gerar Planilha Excel</span>
                        <svg class="w-4 h-4 group-hover:translate-x-0.5 transition-transform" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M14 5l7 7m0 0l-7 7m7-7H3"></path>
                        </svg>
                    </button>
                </form>
            </div>
            <div class="px-8 py-4 bg-slate-50 border-t border-slate-200">
                <p class="text-center text-xs text-slate-500">
                    &copy; <?php echo date('Y'); ?> Conversor RH. Processamento seguro e local.
                </p>
            </div>
        </div>
    </div>
</body>
</html>
