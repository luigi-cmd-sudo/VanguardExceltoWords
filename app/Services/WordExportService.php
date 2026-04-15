<?php

namespace App\Services;

use Illuminate\Support\Collection;
use PhpOffice\PhpWord\TemplateProcessor;
use ZipArchive;

/**
 * WordExportService
 *
 * This service does two things:
 *   1. Generates one .docx file per candidate row
 *   2. Packages all .docx files into one .zip file
 *
 * -----------------------------------------------------------------------
 * FUTURE MODIFICATION GUIDE — Read this before editing this file
 * -----------------------------------------------------------------------
 *
 * TO ADD A NEW PLACEHOLDER:
 *   1. Add ${your_placeholder} to the template via MakeWordTemplate.php
 *      and re-run: php artisan make:word-template
 *   2. In generateOneDocument() below, add a new setValue() line.
 *      Follow the pattern of the existing ones.
 *   3. Make sure your Excel column name matches (see getField() calls).
 *
 * TO CHANGE DATE FORMAT:
 *   Edit the formatDate() method at the bottom of this file.
 *
 * TO CHANGE THE ZIP FILENAME:
 *   Edit the $zipFilename variable in the generateZip() method.
 *
 * TO CHANGE WHERE FILES ARE TEMPORARILY STORED:
 *   Edit the $outputDir variable in generateZip().
 * -----------------------------------------------------------------------
 */
class WordExportService
{
    /**
     * Main entry point.
     * Generates all Word documents and returns the path to the ZIP file.
     *
     * @param  Collection  $rows  All candidate rows from the Excel file
     * @return string             Full path to the generated ZIP file
     */
    public function generateZip(Collection $rows): string
    {
        // --- Check template exists ---
        $templatePath = storage_path('app/private/templates/candidate_template.docx');

        if (!file_exists($templatePath)) {
            throw new \RuntimeException(
                'Template file not found. Please run: php artisan make:word-template'
            );
        }

        // --- Prepare the output folder ---
        // Each run gets its own subfolder using a timestamp.
        // This prevents files from different uploads from mixing together.
        $outputDir = storage_path('app/private/generated/' . now()->format('Ymd_His'));

        if (!is_dir($outputDir)) {
            mkdir($outputDir, 0755, true);
        }

        // --- Generate one Word file per row ---
        $generatedFiles = [];

        foreach ($rows as $row) {
            $docxPath = $this->generateOneDocument($row, $templatePath, $outputDir);
            $generatedFiles[] = $docxPath;
        }

        if (empty($generatedFiles)) {
            throw new \RuntimeException('No documents were generated. Check your Excel file.');
        }

        // --- Create the ZIP file ---
        // -----------------------------------------------------------------------
        // FUTURE MODIFICATION — ZIP filename
        // Change 'generated_documents' below to rename the downloaded ZIP file.
        // -----------------------------------------------------------------------
        $zipFilename = 'generated_documents_' . now()->format('Ymd_His') . '.zip';
        $zipPath     = storage_path('app/private/generated/' . $zipFilename);

        $this->createZip($generatedFiles, $zipPath);

        return $zipPath;
    }

    /**
     * Generates one Word document for one candidate row.
     *
     * @param  mixed   $row          One row from the Excel file
     * @param  string  $templatePath Full path to the .docx template
     * @param  string  $outputDir    Folder where this file will be saved
     * @return string                Full path to the saved .docx file
     */
    private function generateOneDocument(mixed $row, string $templatePath, string $outputDir): string
    {
        // Open the template — a fresh copy for every candidate
        $processor = new TemplateProcessor($templatePath);

        // --- Read values from the Excel row ---
        // getField() safely reads a cell value and returns empty string if missing.
        //
        // -----------------------------------------------------------------------
        // FUTURE MODIFICATION — Adding more Excel columns
        // Copy the pattern below for each new column.
        // The key inside getField() must match your Excel header in snake_case.
        // Example: Excel header "Interview Date" → key 'interview_date'
        // -----------------------------------------------------------------------
        $candidateName       = $this->getField($row, 'candidate_name');
        $caseReferenceNumber = $this->getField($row, 'case_reference_number');
        $caseReceivedDate    = $this->formatDate($this->getField($row, 'case_received_date'));
        $reportDate          = $this->formatDate($this->getField($row, 'report_date'));

        // --- Replace placeholders in the template ---
        // The first argument must match the placeholder name without ${ and }
        // The second argument is the value to insert.
        //
        // -----------------------------------------------------------------------
        // FUTURE MODIFICATION — Adding more placeholders
        // Add one setValue() line per new placeholder.
        // Make sure the placeholder name matches what is in the template.
        // -----------------------------------------------------------------------
        $processor->setValue('candidate_name',       $candidateName);
        $processor->setValue('case_reference_number', $caseReferenceNumber);
        $processor->setValue('case_received_date',   $caseReceivedDate);
        $processor->setValue('report_date',          $reportDate);

        // --- Build a safe filename from the candidate name ---
        // Remove characters that are not allowed in filenames on Windows/Linux.
        // Example: "John/Doe" becomes "JohnDoe.docx"
        //
        // -----------------------------------------------------------------------
        // FUTURE MODIFICATION — File naming format
        // Change the logic below to rename the output files differently.
        // For example, to use case reference as the filename:
        //   $safeFilename = $this->sanitizeFilename($caseReferenceNumber) . '.docx';
        // -----------------------------------------------------------------------
        $safeName     = $this->sanitizeFilename($candidateName ?: 'Unknown_Candidate');
        $safeFilename = $safeName . '.docx';
        $outputPath   = $outputDir . DIRECTORY_SEPARATOR . $safeFilename;

        // If two candidates have the same name, add a number to avoid overwriting
        if (file_exists($outputPath)) {
            $counter    = 2;
            $outputPath = $outputDir . DIRECTORY_SEPARATOR . $safeName . "_{$counter}.docx";
            while (file_exists($outputPath)) {
                $counter++;
                $outputPath = $outputDir . DIRECTORY_SEPARATOR . $safeName . "_{$counter}.docx";
            }
        }

        $processor->saveAs($outputPath);

        return $outputPath;
    }

    /**
     * Creates a ZIP file containing all the generated Word documents.
     *
     * @param  array   $files    Array of full file paths to include
     * @param  string  $zipPath  Full path where the ZIP will be saved
     */
    private function createZip(array $files, string $zipPath): void
    {
        $zip = new ZipArchive();

        // ZipArchive::CREATE creates a new file. ZipArchive::OVERWRITE replaces if exists.
        $result = $zip->open($zipPath, ZipArchive::CREATE | ZipArchive::OVERWRITE);

        if ($result !== true) {
            throw new \RuntimeException(
                'Could not create ZIP file. Check write permissions on storage/app/private/generated/'
            );
        }

        foreach ($files as $filePath) {
            if (file_exists($filePath)) {
                // basename() gives just the filename (e.g. "John Doe.docx")
                // so the ZIP contains flat files, not nested folder paths
                $zip->addFile($filePath, basename($filePath));
            }
        }

        $zip->close();

        if (!file_exists($zipPath)) {
            throw new \RuntimeException('ZIP file was not created successfully.');
        }
    }

    /**
     * Safely reads one cell value from an Excel row.
     * Returns an empty string if the column does not exist.
     *
     * @param  mixed   $row  One Excel row
     * @param  string  $key  Column name in snake_case (matches Excel header)
     * @return string
     */
    private function getField(mixed $row, string $key): string
    {
        $value = $row[$key] ?? '';
        return trim((string) $value);
    }

    /**
     * Formats a date value into "14 APRIL 2026" format.
     *
     * Excel dates can arrive as:
     *   - A number (Excel serial date, e.g. 46125)
     *   - A string (e.g. "2026-04-14" or "04/14/2026")
     *   - Already formatted text
     *
     * -----------------------------------------------------------------------
     * FUTURE MODIFICATION — Date format
     * Change the format string inside date() to use a different format.
     * Current format: 'j F Y' produces "14 April 2026"
     * We then strtoupper() it to get "14 APRIL 2026"
     *
     * Other examples:
     *   'F j, Y'   → "April 14, 2026"
     *   'd/m/Y'    → "14/04/2026"
     *   'Y-m-d'    → "2026-04-14"
     * -----------------------------------------------------------------------
     *
     * @param  string  $value  Raw date value from Excel
     * @return string          Formatted date string, or original value if unparseable
     */
    private function formatDate(string $value): string
    {
        if (empty($value)) {
            return '';
        }

        // Case 1: Excel serial date number (e.g. 46125)
        // Excel counts days from January 1, 1900.
        if (is_numeric($value)) {
            // PHPSpreadsheet/maatwebsite converts Excel serial to Unix timestamp
            $timestamp = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToTimestamp((float) $value);
            return strtoupper(date('j F Y', $timestamp));
        }

        // Case 2: String date — try parsing it
        try {
            $timestamp = strtotime($value);

            if ($timestamp === false) {
                // Could not parse — return original value so document still generates
                return $value;
            }

            return strtoupper(date('j F Y', $timestamp));

        } catch (\Exception $e) {
            return $value;
        }
    }

    /**
     * Removes characters that are not allowed in filenames.
     *
     * Removes: \ / : * ? " < > |
     * Also trims spaces from both ends.
     *
     * @param  string  $name  Raw candidate name
     * @return string         Safe filename (without extension)
     */
    private function sanitizeFilename(string $name): string
    {
        // Replace forbidden characters with nothing
        $safe = preg_replace('/[\\\\\/\:\*\?\"\<\>\|]/', '', $name);

        // Replace multiple spaces with a single space
        $safe = preg_replace('/\s+/', ' ', $safe);

        return trim($safe);
    }
}