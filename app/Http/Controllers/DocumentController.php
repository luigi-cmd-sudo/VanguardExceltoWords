<?php

namespace App\Http\Controllers;

use App\Imports\CandidatesImport;
use App\Services\WordExportService;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Storage;
use Maatwebsite\Excel\Facades\Excel;

class DocumentController extends Controller
{
    // Maximum number of data rows allowed per upload.
    // Row 1 (header) is not counted.
    // -----------------------------------------------------------------------
    // FUTURE MODIFICATION — Row limit
    // Change this number to allow more rows per upload.
    // Be careful: more rows = more memory and processing time.
    // For 500+ rows, consider switching to queued/async processing.
    // -----------------------------------------------------------------------
    private const MAX_ROWS = 200;

    public function __construct(
        private WordExportService $wordService
    ) {}

    // =========================================================================
    // SHOW THE UPLOAD FORM
    // =========================================================================

    public function index()
    {
        return view('upload');
    }

    // =========================================================================
    // HANDLE UPLOAD AND GENERATE DOCUMENTS
    // =========================================================================

    public function upload(Request $request)
    {
        // --- Validate the uploaded file ---
        // -----------------------------------------------------------------------
        // FUTURE MODIFICATION — Validation rules
        // Change 'max:10240' to allow larger files (value is in kilobytes).
        // Change 'mimes:xlsx,xls,csv' to allow other file types.
        // -----------------------------------------------------------------------
        $request->validate([
            'excel_file' => [
                'required',
                'file',
                'mimes:xlsx,xls,csv',
                'max:10240',  // 10 MB maximum
            ],
        ], [
            'excel_file.required' => 'Please select an Excel file.',
            'excel_file.file'     => 'The upload must be a file.',
            'excel_file.mimes'    => 'Only .xlsx, .xls, and .csv files are accepted.',
            'excel_file.max'      => 'The file must not be larger than 10 MB.',
        ]);

        // --- Save the uploaded Excel file temporarily ---
        $excelPath = $request->file('excel_file')->store('uploads');
        $fullPath  = Storage::path($excelPath);

        // --- Read the Excel file ---
        try {
            $import = new CandidatesImport();
            Excel::import($import, $fullPath);
            $rows = $import->getRows();

        } catch (\Exception $e) {
            // Clean up the uploaded file if reading failed
            Storage::delete($excelPath);
            return redirect()->route('upload.form')
                ->withErrors(['excel_file' => 'Could not read the Excel file: ' . $e->getMessage()]);
        }

        // --- Check for empty file ---
        if ($rows->isEmpty()) {
            Storage::delete($excelPath);
            return redirect()->route('upload.form')
                ->withErrors(['excel_file' => 'The Excel file has no data rows. Please check the file.']);
        }

        // --- Check row limit ---
        if ($rows->count() > self::MAX_ROWS) {
            Storage::delete($excelPath);
            return redirect()->route('upload.form')
                ->withErrors([
                    'excel_file' => "The file has {$rows->count()} rows. Maximum allowed is " . self::MAX_ROWS . " rows per upload."
                ]);
        }

        // --- Generate all Word documents and ZIP them ---
        try {
            $zipPath = $this->wordService->generateZip($rows);

        } catch (\RuntimeException $e) {
            Storage::delete($excelPath);
            return redirect()->route('upload.form')
                ->withErrors(['excel_file' => 'Generation failed: ' . $e->getMessage()]);
        }

        //storage path debugging

        //         dd([
        //     'zipPath'          => $zipPath,
        //     'file_exists'      => file_exists($zipPath),
        //     'storage_path'     => storage_path('app/private'),
        //     'DIRECTORY_SEP'    => DIRECTORY_SEPARATOR,
        //     'relativeZipPath'  => str_replace(
        //         storage_path('app/private') . DIRECTORY_SEPARATOR,
        //         '',
        //         $zipPath
        //     ),
        // ]);


        // --- Clean up the uploaded Excel file (no longer needed) ---
        Storage::delete($excelPath);

        // --- Pass the ZIP path to the result page ---
        // We store only the relative path (after storage/app/private/)
        // so we can pass it through the session safely.
        // Normalize all slashes to forward slashes to avoid Windows path issues
        $normalizedZipPath     = str_replace('\\', '/', $zipPath);
        $normalizedStoragePath = str_replace('\\', '/', storage_path('app/private'));

        $relativeZipPath = str_replace(
            $normalizedStoragePath . '/',
            '',
            $normalizedZipPath
        );

        // Store ZIP path and row count in session, then redirect to result page.
        // We use session() instead of passing variables directly because
        // after a redirect the request is brand new — session is the safe way
        // to carry data across a redirect.
        return redirect()->route('upload.result')
            ->with('zip_path',  $relativeZipPath)
            ->with('row_count', $rows->count());
    }

    // =========================================================================
    // SHOW THE RESULT PAGE
    // =========================================================================

    public function result()
    {
        // If someone visits /result directly without uploading, send them back
        if (!session()->has('zip_path')) {
            return redirect()->route('upload.form');
        }

        return view('result', [
            'zipPath'  => session('zip_path'),
            'rowCount' => session('row_count'),
        ]);
    }

    // =========================================================================
    // HANDLE THE ZIP DOWNLOAD
    // =========================================================================

    public function download(Request $request)
    {
        $relativeZipPath = $request->query('file');

        if (empty($relativeZipPath) || str_contains($relativeZipPath, '..')) {
            abort(400, 'Invalid file path.');
        }

        // Normalize slashes for Windows compatibility
        $relativeZipPath = str_replace(['/', '\\'], DIRECTORY_SEPARATOR, $relativeZipPath);

        $fullPath = storage_path('app' . DIRECTORY_SEPARATOR . 'private' . DIRECTORY_SEPARATOR . $relativeZipPath);

        if (!file_exists($fullPath)) {
            return redirect()->route('upload.form')
                ->withErrors(['excel_file' => 'The ZIP file could not be found. Please generate again.']);
        }

        // FIRST: Clean up the individual .docx folders
        // These are the temporary Word documents that were zipped
        // We can delete them now because they are already inside the ZIP
        $generatedDir = storage_path('app/private/generated');
        $folders = glob($generatedDir . '/202*', GLOB_ONLYDIR);
        foreach ($folders as $folder) {
            $files = glob($folder . '/*');
            foreach ($files as $file) {
                if (is_file($file)) {
                    unlink($file);
                }
            }
            rmdir($folder);
        }

        // SECOND: Send the ZIP to the user, then delete the ZIP itself
        return response()->download(
            $fullPath,
            'generated_documents.zip',
            ['Content-Type' => 'application/zip']
        )->deleteFileAfterSend(true);
    }
}