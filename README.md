markdown
# 📄 Excel to Word Converter

A simple Laravel 11 web application that converts Excel spreadsheet rows into individual Word documents, then packages them all into a single ZIP file for easy download.

**No database. No queues. No background jobs.** Just upload → process → download.

---

## 📖 Table of Contents

- [What Does This App Do?](#what-does-this-app-do)
- [Requirements](#requirements)
- [How to Set Up This Project on Your Computer](#how-to-set-up-this-project-on-your-computer)
  - [Option A — Cloning from GitHub](#option-a--cloning-from-github)
  - [Option B — Starting from Scratch](#option-b--starting-from-scratch)
- [How to Run the Application](#how-to-run-the-application)
- [How to Prepare Your Excel File](#how-to-prepare-your-excel-file)
- [Project Structure — What Each File Does](#project-structure--what-each-file-does)
- [How to Modify the Code](#how-to-modify-the-code)
  - [Safe to Edit — Go Ahead and Change These](#-safe-to-edit--go-ahead-and-change-these)
  - [Do NOT Touch — Leave These Alone](#-do-not-touch--leave-these-alone)
  - [Common Modifications Step by Step](#common-modifications-step-by-step)
- [Quick Reference Table](#-quick-reference-table)
- [Troubleshooting](#-troubleshooting)
- [Limitations](#-limitations)
- [License](#-license)

---

## What Does This App Do?

1. You upload an Excel file (`.xlsx`, `.xls`, or `.csv`) containing candidate data.
2. The app reads every row in the file.
3. For each row, one Word document (`.docx`) is created using a template.
4. All Word documents are bundled into a single `.zip` file.
5. You download the ZIP file — done!

---

## Requirements

Before you begin, make sure the following are installed on your computer:

|
 Software 
|
 Minimum Version 
|
 How to Check 
|
|
----------
|
----------------
|
--------------
|
|
**
PHP
**
|
 8.2 or higher 
|
 Run 
`php -v`
 in your terminal 
|
|
**
Composer
**
|
 2.x 
|
 Run 
`composer -V`
 in your terminal 
|
|
**
Laravel
**
|
 11.x 
|
 Included when you install the project 
|
|
**
Git
**
|
 Any recent version 
|
 Run 
`git --version`
 in your terminal 
|

> **Not sure how to install these?**
> - PHP & Composer: Download [XAMPP](https://www.apachefriends.org/) (includes PHP) and [Composer](https://getcomposer.org/download/).
> - Git: Download from [git-scm.com](https://git-scm.com/downloads).

---

## How to Set Up This Project on Your Computer

### Option A — Cloning from GitHub

Use this if someone has already shared the project on GitHub.

**Step 1 — Clone the repository**

Open your terminal (Command Prompt, PowerShell, or Terminal on Mac) and run:

```bash
git clone https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
cd YOUR_REPO_NAME
Replace YOUR_USERNAME/YOUR_REPO_NAME with the actual GitHub URL.

Step 2 — Install PHP dependencies

bash
composer install
This downloads all the packages the project needs. It may take a minute or two.

Step 3 — Create the environment file

bash
cp .env.example .env
On Windows (if cp does not work):

bash
copy .env.example .env
Step 4 — Generate the application key

bash
php artisan key:generate
Step 5 — Create the required storage folders

On Mac/Linux:

bash
mkdir -p storage/app/private/templates
mkdir -p storage/app/private/generated
On Windows:

bash
mkdir storage\app\private\templates
mkdir storage\app\private\generated
Step 6 — Create the Word template file

bash
php artisan make:word-template
You should see: Template created at: .../storage/app/private/templates/candidate_template.docx

You only need to run this once. The template file will be saved inside your project.

Step 7 — Update your PHP settings

Find your php.ini file. To locate it, run:

bash
php --ini
Open the file and make sure these lines are set (search for each one and change the value):

ini
upload_max_filesize = 10M
post_max_size = 12M
memory_limit = 256M
max_execution_time = 120
After changing php.ini, restart your web server (or restart XAMPP) for the changes to take effect.

Step 8 — You're ready! Jump to How to Run the Application.

Option B — Starting from Scratch
Use this if you want to build the project yourself from zero.

Step 1 — Create a new Laravel project

bash
composer create-project laravel/laravel excel-to-word
cd excel-to-word
Step 2 — Install the required packages

bash
composer require maatwebsite/excel
composer require phpoffice/phpword
Step 3 — Publish the Excel configuration

bash
php artisan vendor:publish --provider="Maatwebsite\Excel\ExcelServiceProvider" --tag=config
Step 4 — Create the storage folders

On Mac/Linux:

bash
mkdir -p storage/app/private/templates
mkdir -p storage/app/private/generated
On Windows:

bash
mkdir storage\app\private\templates
mkdir storage\app\private\generated
Step 5 — Create all the project files

You need to create the following files manually. The full code for each file is provided in the Source Code Reference section below.

File	What to Do
app/Console/Commands/MakeWordTemplate.php	Run php artisan make:command MakeWordTemplate, then paste the code
app/Imports/CandidatesImport.php	Create the file manually and paste the code
app/Services/WordExportService.php	Create the Services folder and file, then paste the code
app/Http/Controllers/DocumentController.php	Run php artisan make:controller DocumentController, then paste the code
routes/web.php	Replace the contents with the route definitions
resources/views/upload.blade.php	Create the file and paste the code
resources/views/result.blade.php	Create the file and paste the code
Step 6 — Generate the Word template

bash
php artisan make:word-template
Step 7 — Update your PHP settings (same as Option A, Step 7 above)

Step 8 — You're ready! Jump to How to Run the Application.

How to Run the Application
Step 1 — Start the development server

bash
php artisan serve
Step 2 — Open your browser and go to:

text
http://127.0.0.1:8000
Step 3 — Upload your Excel file, wait for processing, and download the ZIP.

That's it! 🎉

How to Prepare Your Excel File
Your Excel file must follow these rules:

Row 1 must be the header row — these are your column names.
The following column headers are required (spelled exactly like this):
Column Header	Example Value
Candidate Name	John Doe
Case Reference Number	REF-2026-001
Case Received Date	2026-01-15
Report Date	2026-04-14
Dates can be in Excel date format or typed as text (e.g., 2026-04-14 or 04/14/2026).
Maximum 200 data rows per upload (not counting the header).
Empty rows at the bottom are automatically skipped.
Sample Excel File
Candidate Name	Case Reference Number	Case Received Date	Report Date
John Doe	REF-2026-001	2026-01-15	2026-04-14
Maria Santos	REF-2026-002	2026-02-20	2026-04-14
James Cruz	REF-2026-003	2026-03-10	2026-04-14
Project Structure — What Each File Does
text
your-project/
│
├── app/
│   ├── Console/
│   │   └── Commands/
│   │       └── MakeWordTemplate.php        ← Creates the Word template file
│   │
│   ├── Http/
│   │   └── Controllers/
│   │       └── DocumentController.php      ← Handles upload, processing, and download
│   │
│   ├── Imports/
│   │   └── CandidatesImport.php            ← Reads the Excel file into rows
│   │
│   └── Services/
│       └── WordExportService.php           ← Generates Word files and creates the ZIP
│
├── resources/
│   └── views/
│       ├── upload.blade.php                ← The upload form page (what users see first)
│       └── result.blade.php                ← The download page (shown after processing)
│
├── routes/
│   └── web.php                             ← Defines the 4 URLs the app uses
│
└── storage/
    └── app/
        └── private/
            ├── templates/
            │   └── candidate_template.docx ← The Word template (auto-generated)
            └── generated/                  ← Where ZIP files are temporarily stored
What Each File Does (Plain English)
File	Purpose
MakeWordTemplate.php	A command you run once. It creates the Word document template with placeholders like ${candidate_name} that get replaced with real data.
CandidatesImport.php	Reads your Excel file. It turns each row into data that the app can use. You almost never need to touch this file.
WordExportService.php	The heart of the app. It takes each candidate's data, fills in the Word template, saves the file, and then puts all the files into a ZIP.
DocumentController.php	The traffic controller. It receives the uploaded file, calls the other files to do their jobs, handles errors, and sends the ZIP back to you.
web.php	Tells Laravel which URLs exist and which controller method to call for each one.
upload.blade.php	The HTML page with the file upload form.
result.blade.php	The HTML page with the download button that appears after documents are generated.
How to Modify the Code
✅ Safe to Edit — Go Ahead and Change These
These files are designed to be customized. Editing them will not break the app (as long as you follow the patterns already there).

File	What You Can Change
app/Console/Commands/MakeWordTemplate.php	Word document layout, fonts, colors, margins, placeholders, header text
app/Services/WordExportService.php	Date format, output filename per document, ZIP filename, adding new Excel column mappings
app/Http/Controllers/DocumentController.php	Row limit (max 200), file size limit, validation error messages
resources/views/upload.blade.php	Upload page design — colors, text, layout, CSS
resources/views/result.blade.php	Result page design — button text, colors, layout, CSS
🚫 Do NOT Touch — Leave These Alone
Unless you are an experienced developer, do not edit these files. Changing them can break the app in ways that are hard to debug.

File	Why You Should Leave It Alone
app/Imports/CandidatesImport.php	It automatically reads ALL Excel columns. You never need to change it when adding new columns.
routes/web.php	The four routes are all the app needs. Adding or changing routes without understanding Laravel routing can break navigation.
.env	Only change this if you know what each setting does. The defaults work fine for local development.
config/excel.php	Auto-generated config for the Excel package. No need to change it.
composer.json / composer.lock	Managed by Composer. Do not edit manually.
Any file inside vendor/	These are third-party packages. Never edit anything in the vendor folder. Your changes will be lost when you run composer install.
Common Modifications Step by Step
🎨 Change the Word Document Design
Want to change fonts, colors, margins, or the text layout in the generated Word files?

File to edit: app/Console/Commands/MakeWordTemplate.php

Open the file.
Find the handle() method.
Edit the fonts, colors, spacing, or text as needed.
Save the file.
Run this command to rebuild the template:
bash
php artisan make:word-template
Examples of what you can change:

What	Where in the Code	Example
Page margins	$phpWord->addSection([...])	Change 1440 (1 inch) to 720 (half inch)
Default font	setDefaultFontName('Calibri')	Change 'Calibri' to 'Arial' or 'Times New Roman'
Header text	$header->addText(...)	Change the string to whatever you want
Label colors	'color' => '2C3E50'	Replace with any hex color code (without the #)
➕ Add a New Excel Column and Word Placeholder
Want to add a new field like "Interview Date" to your documents?

You need to edit two files and update your Excel file.

File 1 — app/Console/Commands/MakeWordTemplate.php

Add these two lines inside the handle() method, after the existing placeholders:

php
$section->addText('Interview Date:', $labelFont, $spacing);
$section->addText('${interview_date}', $valueFont, ['spaceAfter' => 320]);
Then rebuild the template:

bash
php artisan make:word-template
File 2 — app/Services/WordExportService.php

In the generateOneDocument() method, add:

php
// Read the new column
$interviewDate = $this->formatDate($this->getField($row, 'interview_date'));

// Replace the placeholder in the template
$processor->setValue('interview_date', $interviewDate);
File 3 — Your Excel file

Add a new column called Interview Date in Row 1 (the header row).

Important: The Excel column header Interview Date automatically becomes interview_date (lowercase, spaces become underscores). Always use the snake_case version in your code.

📅 Change the Date Format
Want dates to appear as 04/14/2026 instead of 14 APRIL 2026?

File to edit: app/Services/WordExportService.php

Find the formatDate() method. Look for this line:

php
return strtoupper(date('j F Y', $timestamp));
Replace it with one of these formats:

Format Code	Output Example
'j F Y' (current)	14 APRIL 2026
'd/m/Y'	14/04/2026
'm/d/Y'	04/14/2026
'F j, Y'	April 14, 2026
'Y-m-d'	2026-04-14
Note: There are two return lines with date() in that method (one for Excel serial dates, one for text dates). Change both of them.

📁 Change the Generated Filename for Each Document
Want files named by case reference instead of candidate name?

File to edit: app/Services/WordExportService.php

Find this line in generateOneDocument():

php
$safeName = $this->sanitizeFilename($candidateName ?: 'Unknown_Candidate');
Replace it with:

php
// Use case reference number as filename
$safeName = $this->sanitizeFilename($caseReferenceNumber ?: 'Unknown_Reference');
Or combine both:

php
// Use name + reference as filename
$safeName = $this->sanitizeFilename($candidateName . '_' . $caseReferenceNumber);
⬆️ Increase the Row Limit
Need to process more than 200 rows at once?

File to edit: app/Http/Controllers/DocumentController.php

Find this line near the top of the file:

php
private const MAX_ROWS = 200;
Change 200 to your desired limit.

⚠️ Warning: Going above 500 rows may cause the server to time out or run out of memory. For very large files (500+ rows), consider upgrading to queue-based processing — that is a more advanced topic not covered here.

🗜️ Change the Downloaded ZIP Filename
Want the ZIP file to have a different name when the user downloads it?

File to edit: app/Http/Controllers/DocumentController.php

Find this line in the download() method:

php
return response()->download($fullPath, 'generated_documents.zip', [...]);
Change 'generated_documents.zip' to whatever name you prefer:

php
return response()->download($fullPath, 'candidate_reports.zip', [...]);
🖼️ Change the Upload or Result Page Design
Want to change colors, button text, or page layout?

What to change	File to edit
Upload page	resources/views/upload.blade.php
Result/download page	resources/views/result.blade.php
Both files contain a <style> section at the top for CSS (colors, fonts, spacing) and HTML below it for the page content. Edit them just like any normal HTML file.

📋 Quick Reference Table
I want to change...	Edit this file
Word document design/layout	app/Console/Commands/MakeWordTemplate.php
Page header text in the Word doc	app/Console/Commands/MakeWordTemplate.php
Add/remove Word placeholders	MakeWordTemplate.php + WordExportService.php
How Excel columns map to placeholders	app/Services/WordExportService.php
Date format	app/Services/WordExportService.php → formatDate()
Output filename per candidate	app/Services/WordExportService.php → generateOneDocument()
ZIP filename	WordExportService.php + DocumentController.php
Upload form design	resources/views/upload.blade.php
Result page / button design	resources/views/result.blade.php
Row limit	app/Http/Controllers/DocumentController.php
File size limit	app/Http/Controllers/DocumentController.php
Download behavior	app/Http/Controllers/DocumentController.php → download()
🔧 Troubleshooting
"The ZIP file could not be found. Please generate again."
This is usually a Windows path issue where forward slashes and backslashes get mixed up.

Fix: Open app/Http/Controllers/DocumentController.php.

Find the str_replace block in the upload() method (around line 100) and replace it with:

php
// Normalize all slashes to forward slashes to avoid Windows path issues
$normalizedZipPath     = str_replace('\\', '/', $zipPath);
$normalizedStoragePath = str_replace('\\', '/', storage_path('app/private'));

$relativeZipPath = str_replace(
    $normalizedStoragePath . '/',
    '',
    $normalizedZipPath
);
Also update the download() method. Add this line right after reading the query parameter:

php
$relativeZipPath = str_replace(['/', '\\'], DIRECTORY_SEPARATOR, $relativeZipPath);
And update the full path construction to:

php
$fullPath = storage_path('app' . DIRECTORY_SEPARATOR . 'private' . DIRECTORY_SEPARATOR . $relativeZipPath);
"Template file not found"
You forgot to generate the Word template. Run:

bash
php artisan make:word-template
"Could not read the Excel file"
Make sure your file is .xlsx, .xls, or .csv format.
Make sure Row 1 contains the column headers.
Make sure the file is not open in Excel while uploading.
"The file has X rows. Maximum allowed is 200 rows per upload."
Your file has too many rows. Either:

Split your Excel file into smaller files (200 rows or fewer each), or
Increase the row limit (see Increase the Row Limit).
Page loads but nothing happens / blank screen
Make sure the Laravel server is running: php artisan serve
Check the terminal where php artisan serve is running for error messages.
Check storage/logs/laravel.log for detailed error information.
⚠️ Limitations
Maximum 200 rows per upload (configurable, but going too high may cause timeouts).
Maximum 10 MB file size (configurable in the controller).
No database — nothing is saved permanently. Generated files are temporary.
No user accounts — anyone who can access the URL can use the app.
Synchronous processing only — the user must wait while files are generated.
For very large batches (500+ rows), consider upgrading to Laravel's queue system for background processing.
