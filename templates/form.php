<!DOCTYPE html>
<html lang="es" class="h-full bg-slate-950">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Conversor de Marcaciones</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Inter', sans-serif; }
    </style>
</head>
<body class="h-full flex flex-col items-center justify-center p-6 text-slate-100">
    <div class="w-full max-w-md">
        <div class="bg-slate-900 rounded-xl shadow-2xl border border-slate-800 overflow-hidden">
            <div class="p-8">
                <div class="flex flex-col items-center mb-8">
                    <div class="w-12 h-12 bg-indigo-600 rounded-lg flex items-center justify-center mb-4 shadow-lg shadow-indigo-500/20" style="width: 3rem; height: 3rem;">
                        <svg class="w-6 h-6 text-white" style="width: 1.5rem; height: 1.5rem;" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
                        </svg>
                    </div>
                    <h1 class="text-2xl font-bold text-white tracking-tight text-center">Conversor de Marcaciones</h1>
                    <p class="text-slate-400 mt-2 text-center text-sm">Convierta sus archivos TXT en planillas Excel formateadas.</p>
                </div>

                <?php if (!empty($error)): ?>
                    <div class="mb-6 p-4 bg-red-900/30 border border-red-800 rounded-lg flex items-start gap-3">
                        <svg class="w-5 h-5 text-red-500 mt-0.5 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path>
                        </svg>
                        <p class="text-sm text-red-200 font-medium"><?php echo htmlspecialchars($error); ?></p>
                    </div>
                <?php endif; ?>

                <form action="" method="post" enctype="multipart/form-data" class="space-y-6">
                    <input type="hidden" name="csrf_token" value="<?php echo htmlspecialchars($csrfToken); ?>">
                    <div>
                        <label for="file" class="block text-sm font-semibold text-slate-300 mb-2">Seleccione el archivo de marcaciones</label>
                        <div class="relative group">
                            <input type="file" name="file" id="file" required
                                class="block w-full text-sm text-slate-400
                                file:mr-4 file:py-2.5 file:px-4
                                file:rounded-lg file:border-0
                                file:text-sm file:font-semibold
                                file:bg-indigo-900/50 file:text-indigo-300
                                hover:file:bg-indigo-900/70
                                border border-slate-700 bg-slate-800 rounded-lg p-1.5
                                group-hover:border-indigo-500 transition-colors
                                focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500">
                        </div>
                        <p class="mt-2 text-xs text-slate-500 italic">Formatos compatibles: TXT (máx. 1MB).</p>
                    </div>

                    <button type="submit"
                        class="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-semibold py-3 px-4 rounded-lg shadow-lg shadow-indigo-500/20 transition-all duration-200 flex items-center justify-center gap-2 group focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:ring-offset-2 focus:ring-offset-slate-900">
                        <span>Generar Planilla Excel</span>
                            <svg class="w-4 h-4 group-hover:translate-x-0.5 transition-transform" style="width: 1rem; height: 1rem;" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M14 5l7 7m0 0l-7 7m7-7H3"></path>
                        </svg>
                    </button>
                </form>
            </div>
            <div class="px-8 py-4 bg-slate-900/50 border-t border-slate-800">
                <p class="text-center text-xs text-slate-500">
                    &copy; <?php echo date('Y'); ?> Conversor RH. Procesamiento seguro e local.
                </p>
            </div>
        </div>
    </div>
</body>
</html>
