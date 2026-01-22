<?php

namespace App\Service;

use League\Csv\Reader;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

class ConversorService
{
    public function convertToSpreadsheet(string $filePath): Spreadsheet
    {
        $csv = Reader::createFromPath($filePath, 'r');
        $csv->setDelimiter("\t");
        $csv->setHeaderOffset(0);

        $headers = $csv->getHeader();
        $records = $csv->getRecords();
        $collaborators = [];

        // Mapping for headers
        $map = [
            'dept' => ['Departamento', 'Dept.'],
            'id' => ['ID Colaborador', 'User ID', 'Enroll ID', 'ID'],
            'name' => ['Nome', 'Name', 'Nombre'],
            'date' => ['Data', 'Date', 'Fecha'],
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

            if ($hM1 || $hM2 || $hM3 || $hM4) {
                $marks = [];
                if (!empty($record[$hM1])) $marks[] = $record[$hM1];
                if (!empty($record[$hM2])) $marks[] = $record[$hM2];
                if (!empty($record[$hM3])) $marks[] = $record[$hM3];
                if (!empty($record[$hM4])) $marks[] = $record[$hM4];
                $collaborators[$id]['days'][$date] = array_slice($marks, 0, 4);
            } else {
                $time = $record[$hTime] ?? '';
                if (!empty($time) && count($collaborators[$id]['days'][$date]) < 4) {
                    $collaborators[$id]['days'][$date][] = $time;
                }
            }
        }

        if (empty($collaborators)) {
            throw new \Exception('Archivo vacío o en formato inválido.');
        }

        $spreadsheet = new Spreadsheet();
        $spreadsheet->removeSheetByIndex(0);

        foreach ($collaborators as $id => $data) {
            $name = $data['name'];
            $sheetName = mb_substr(preg_replace('/[\*\?\:\\\\\/\[\]]/', '', $name), 0, 31);
            if (empty($sheetName)) {
                $sheetName = "Colaborador $id";
            }

            $baseSheetName = $sheetName;
            $counter = 1;
            while ($spreadsheet->sheetNameExists($sheetName)) {
                $sheetName = mb_substr($baseSheetName, 0, 31 - strlen((string)$counter) - 1) . "($counter)";
                $counter++;
            }

            $sheet = $spreadsheet->createSheet();
            $sheet->setTitle($sheetName);

            $sheet->setCellValue('A1', 'Nombre');
            $sheet->setCellValue('B1', 'Fecha');
            $sheet->setCellValue('C1', '1');
            $sheet->setCellValue('D1', '2');
            $sheet->setCellValue('E1', '3');
            $sheet->setCellValue('F1', '4');
            $sheet->setCellValue('G1', 'Mañana');
            $sheet->setCellValue('H1', 'Tarde');
            $sheet->setCellValue('I1', 'Total');

            $sheet->getStyle('A1:I1')->getFont()->setBold(true);

            $row = 2;
            $startRow = 2;
            foreach ($data['days'] as $date => $times) {
                sort($times);
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

                $sheet->setCellValue("G$row", "=IF(AND(C$row<>\"\",D$row<>\"\"),D$row-C$row,\"\")");
                $sheet->setCellValue("H$row", "=IF(AND(E$row<>\"\",F$row<>\"\"),F$row-E$row,\"\")");
                $sheet->setCellValue("I$row", "=IF(AND(G$row=\"\",H$row=\"\"),\"\",IF(G$row=\"\",0,G$row)+IF(H$row=\"\",0,H$row))");

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
            $row++;
            $sheet->setCellValue("I$row", "=SUM(I$startRow:I$endDataRow)");
            $sheet->getStyle("I$row")->getNumberFormat()->setFormatCode('[h]:mm:ss');
            $sheet->getStyle("I$row")->getFont()->setBold(true);

            foreach (range('A', 'I') as $col) {
                $sheet->getColumnDimension($col)->setAutoSize(true);
            }
        }

        return $spreadsheet;
    }
}
