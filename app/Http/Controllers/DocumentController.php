<?php

namespace App\Http\Controllers;

use App\Imports\CandidatesImport;
use App\Services\WordExportService;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Storage;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Support\Facades\File;

class DocumentController extends Controller
{
    private const MAX_ROWS = 200;

    // -----------------------------------------------------------------------
    // Files older than this (in minutes) will be auto-deleted
    // Change this number if you want files to live longer or shorter
    // -----------------------------------------------------------------------
    private const CLEANUP_AFTER_MINUTES = 60;

    public function __construct(
        private WordExportService $wordService
    ) {}

    // =========================================================================
    // SHOW THE UPLOAD FORM
    // =========================================================================

    public function index()
    {
        // Clean up old files every time someone visits the upload page
        $this->cleanupOldFiles();

        return view('upload');
    }

    // =========================================================================
    // HANDLE UPLOAD AND GENERATE DOCUMENTS
    // =========================================================================

    public function upload(Request $request)
    {
        // Clean up old files before generating new ones
        $this->cleanupOldFiles();

        $request->validate([
            'excel_file' => [
                'required',
                'file',
                'mimes:xlsx,xls,csv',
                'max:10240',
            ],
        ], [
            'excel_file.required' => 'Please select an Excel file.',
            'excel_file.file'     => 'The upload must be a file.',
            'excel_file.mimes'    => 'Only .xlsx, .xls, and .csv files are accepted.',
            'excel_file.max'      => 'The file must not be larger than 10 MB.',
        ]);

        $excelPath = $request->file('excel_file')->store('uploads');
        $fullPath  = Storage::path($excelPath);

        try {
            $import = new CandidatesImport();
            Excel::import($import, $fullPath);
            $rows = $import->getRows();

        } catch (\Exception $e) {
            Storage::delete($excelPath);
            return redirect()->route('upload.form')
                ->withErrors(['excel_file' => 'Could not read the Excel file: ' . $e->getMessage()]);
        }

        if ($rows->isEmpty()) {
            Storage::delete($excelPath);
            return redirect()->route('upload.form')
                ->withErrors(['excel_file' => 'The Excel file has no data rows. Please check the file.']);
        }

        if ($rows->count() > self::MAX_ROWS) {
            Storage::delete($excelPath);
            return redirect()->route('upload.form')
                ->withErrors([
                    'excel_file' => "The file has {$rows->count()} rows. Maximum allowed is " . self::MAX_ROWS . " rows per upload."
                ]);
        }

        try {
            $zipPath = $this->wordService->generateZip($rows);

        } catch (\RuntimeException $e) {
            Storage::delete($excelPath);
            return redirect()->route('upload.form')
                ->withErrors(['excel_file' => 'Generation failed: ' . $e->getMessage()]);
        }

        Storage::delete($excelPath);

        $normalizedZipPath     = str_replace('\\', '/', $zipPath);
        $normalizedStoragePath = str_replace('\\', '/', storage_path('app/private'));

        $relativeZipPath = str_replace(
            $normalizedStoragePath . '/',
            '',
            $normalizedZipPath
        );

        return redirect()->route('upload.result')
            ->with('zip_path',  $relativeZipPath)
            ->with('row_count', $rows->count());
    }

    // =========================================================================
    // SHOW THE RESULT PAGE
    // =========================================================================

    public function result()
    {
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

        $relativeZipPath = str_replace(['/', '\\'], DIRECTORY_SEPARATOR, $relativeZipPath);

        $fullPath = storage_path('app' . DIRECTORY_SEPARATOR . 'private' . DIRECTORY_SEPARATOR . $relativeZipPath);

        if (!file_exists($fullPath)) {
            return redirect()->route('upload.form')
                ->withErrors(['excel_file' => 'The ZIP file could not be found. Please generate again.']);
        }

        // Clean up the individual .docx folders (already inside the ZIP)
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

        // Send the ZIP then delete it
        return response()->download(
            $fullPath,
            'generated_documents.zip',
            ['Content-Type' => 'application/zip']
        )->deleteFileAfterSend(true);
    }

    // =========================================================================
    // AUTO-CLEANUP OLD FILES
    // =========================================================================

    /**
     * Deletes all generated files older than CLEANUP_AFTER_MINUTES.
     *
     * This runs every time someone:
     *   - Visits the upload page
     *   - Uploads a new file
     *
     * This prevents storage from filling up when users don't download.
     */
    private function cleanupOldFiles(): void
    {
        $generatedDir = storage_path('app/private/generated');

        // If the folder doesn't exist, nothing to clean
        if (!is_dir($generatedDir)) {
            return;
        }

        $now = time();
        $maxAge = self::CLEANUP_AFTER_MINUTES * 30; // Convert minutes to seconds

        // --- Clean up old date folders (e.g., 20260415_020852/) ---
        $folders = glob($generatedDir . '/202*', GLOB_ONLYDIR);
        foreach ($folders as $folder) {
            if (($now - filemtime($folder)) > $maxAge) {
                // Delete all files inside the folder
                $files = glob($folder . '/*');
                foreach ($files as $file) {
                    if (is_file($file)) {
                        unlink($file);
                    }
                }
                // Delete the empty folder
                rmdir($folder);
            }
        }

        // --- Clean up old ZIP files ---
        $zipFiles = glob($generatedDir . '/*.zip');
        foreach ($zipFiles as $zipFile) {
            if (($now - filemtime($zipFile)) > $maxAge) {
                unlink($zipFile);
            }
        }

        // --- Clean up old uploaded Excel files ---
        $uploadsDir = storage_path('app/private/uploads');
        if (is_dir($uploadsDir)) {
            $uploadedFiles = glob($uploadsDir . '/*');
            foreach ($uploadedFiles as $uploadedFile) {
                if (is_file($uploadedFile) && ($now - filemtime($uploadedFile)) > $maxAge) {
                    unlink($uploadedFile);
                }
            }
        }
    }
}