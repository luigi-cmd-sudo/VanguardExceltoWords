<?php

namespace App\Imports;

use Illuminate\Support\Collection;
use Maatwebsite\Excel\Concerns\ToCollection;
use Maatwebsite\Excel\Concerns\WithHeadingRow;

/**
 * CandidatesImport
 *
 * Reads the Excel file and returns all rows as a Collection.
 *
 * WithHeadingRow means row 1 is treated as column headers.
 * The headers are automatically converted to snake_case with spaces
 * replaced by underscores and all lowercase.
 *
 * So your Excel column "Candidate Name" becomes $row['candidate_name'].
 * And "Case Reference Number" becomes $row['case_reference_number'].
 * And "Case Received Date" becomes $row['case_received_date'].
 * And "Report Date" becomes $row['report_date'].
 *
 * -----------------------------------------------------------------------
 * FUTURE MODIFICATION — Adding Excel Columns
 * -----------------------------------------------------------------------
 * If you add a new column to your Excel file, you do NOT need to change
 * this file. WithHeadingRow reads ALL columns automatically.
 * You only need to update WordExportService.php to use the new column.
 * -----------------------------------------------------------------------
 */
class CandidatesImport implements ToCollection, WithHeadingRow
{
    private Collection $rows;

    public function __construct()
    {
        $this->rows = new Collection();
    }

    public function collection(Collection $rows): void
    {
        // Filter out empty rows (blank rows at the bottom of Excel files)
        $this->rows = $rows->filter(function ($row) {
            return collect($row)->filter(fn($v) => trim((string) $v) !== '')->isNotEmpty();
        });
    }

    public function getRows(): Collection
    {
        return $this->rows;
    }
}