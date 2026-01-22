<?php

require __DIR__ . '/../vendor/autoload.php';

use App\Security\Csrf;
use App\Security\RateLimiter;
use App\Service\ConversorService;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

session_start();

// Security Headers
header("X-Content-Type-Options: nosniff");
header("X-Frame-Options: DENY");
header("Referrer-Policy: no-referrer");
header("Content-Security-Policy: default-src 'self'; script-src 'https://cdn.tailwindcss.com'; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; font-src https://fonts.gstatic.com; img-src 'self' data:;");

$error = '';

try {
    if ($_SERVER['REQUEST_METHOD'] === 'POST') {
        // 1. Rate Limit
        $limiter = new RateLimiter(__DIR__ . '/../var/ratelimit/');
        if (!$limiter->check($_SERVER['REMOTE_ADDR'])) {
            throw new Exception('Demasiadas solicitudes. Intente de nuevo más tarde.');
        }

        // 2. CSRF Validation
        if (!Csrf::validate($_POST['csrf_token'] ?? '')) {
            throw new Exception('Token de seguridad inválido.');
        }

        // 3. File Validation
        if (!isset($_FILES['file']) || $_FILES['file']['error'] !== UPLOAD_ERR_OK) {
            throw new Exception('Error al subir el archivo.');
        }

        $file = $_FILES['file'];

        // Size limit 1MB
        if ($file['size'] > 1024 * 1024) {
            throw new Exception('El archivo es demasiado grande (máx. 1MB).');
        }

        // Check for null bytes
        if (strpos(file_get_contents($file['tmp_name']), "\0") !== false) {
            throw new Exception('El archivo contiene caracteres inválidos.');
        }

        // Validate MIME type
        $finfo = new finfo(FILEINFO_MIME_TYPE);
        $mime = $finfo->file($file['tmp_name']);
        $allowedMimes = ['text/plain', 'text/tsv', 'application/octet-stream']; // octet-stream sometimes for no extension
        if (!in_array($mime, $allowedMimes)) {
            throw new Exception('Tipo de archivo no permitido.');
        }

        // 4. Process
        $service = new ConversorService();
        $spreadsheet = $service->convertToSpreadsheet($file['tmp_name']);

        // 5. Output
        $writer = new Xlsx($spreadsheet);

        // Use a temp stream to get the size
        $tempStream = fopen('php://temp', 'rw+');
        $writer->save($tempStream);
        $size = ftell($tempStream);
        rewind($tempStream);
        $content = stream_get_contents($tempStream);
        fclose($tempStream);

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="marcaciones.xlsx"');
        header('Cache-Control: max-age=0');
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT');
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');
        header('Cache-Control: cache, must-revalidate');
        header('Pragma: public');
        header('Content-Length: ' . $size);

        echo $content;
        exit;
    }
} catch (Exception $e) {
    // Log error
    $logEntry = sprintf("[%s] %s: %s\n", date('Y-m-d H:i:s'), $_SERVER['REMOTE_ADDR'], $e->getMessage());
    file_put_contents(__DIR__ . '/../var/app.log', $logEntry, FILE_APPEND);

    $error = $e->getMessage();
}

$csrfToken = Csrf::generate();
require __DIR__ . '/../templates/form.php';
