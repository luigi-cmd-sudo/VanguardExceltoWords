<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;

class MakeWordTemplate extends Command
{
    protected $signature   = 'make:word-template';
    protected $description = 'Create the candidate_template.docx file';

    public function handle(): int
    {
        $phpWord = new PhpWord();
        $phpWord->setDefaultFontName('Calibri');
        $phpWord->setDefaultFontSize(11);

        // --- Page Setup ---
        // All measurements are in twips. 1 inch = 1440 twips.
        $section = $phpWord->addSection([
            'marginTop'    => 1440,
            'marginBottom' => 1440,
            'marginLeft'   => 1800,
            'marginRight'  => 1800,
        ]);

        // --- Page Header ---
        // ${report_date} will be replaced with the actual date from Excel.
        // To change the header text, edit the string below.
        $header = $section->addHeader();
        $header->addText(
            'CANDIDATE REPORT  |  Date: ${report_date}',
            ['bold' => true, 'size' => 11, 'color' => '2C3E50']
        );

        // --- Document Body ---
        // These are the placeholders that will be replaced with real data.
        // To add more fields: copy one addText() block and change the placeholder.
        // Remember to also update WordExportService.php when you add placeholders.

        $labelFont = ['bold' => true, 'size' => 11, 'color' => '2C3E50'];
        $valueFont = ['size' => 11, 'color' => '333333'];
        $spacing   = ['spaceAfter' => 160];

        $section->addText('CANDIDATE INFORMATION', ['bold' => true, 'size' => 14, 'color' => '2C3E50'], ['spaceAfter' => 240]);
        $section->addText('_____________________________________________', ['color' => 'AAAAAA'], ['spaceAfter' => 240]);

        // Candidate Name
        $section->addText('Candidate Name:', $labelFont, $spacing);
        $section->addText('${candidate_name}', $valueFont, ['spaceAfter' => 320]);

        // Case Reference Number
        $section->addText('Case Reference Number:', $labelFont, $spacing);
        $section->addText('${case_reference_number}', $valueFont, ['spaceAfter' => 320]);

        // Case Received Date
        $section->addText('Case Received Date:', $labelFont, $spacing);
        $section->addText('${case_received_date}', $valueFont, ['spaceAfter' => 320]);

        // Report Date
        $section->addText('Report Date:', $labelFont, $spacing);
        $section->addText('${report_date}', $valueFont, ['spaceAfter' => 320]);

        // --- Save ---
        $savePath = storage_path('app/private/templates/candidate_template.docx');

        if (!is_dir(dirname($savePath))) {
            mkdir(dirname($savePath), 0755, true);
        }

        IOFactory::createWriter($phpWord, 'Word2007')->save($savePath);

        $this->info('Template created at: ' . $savePath);
        return self::SUCCESS;
    }
}