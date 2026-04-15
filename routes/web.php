<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\DocumentController;

/*
|--------------------------------------------------------------------------
| Routes
|--------------------------------------------------------------------------
|
| GET  /          → Show the upload form
| POST /upload    → Process the Excel file and generate documents
| GET  /result    → Show the result page with download buttons
| GET  /download  → Download the ZIP file
|
*/

// Upload form
Route::get('/', [DocumentController::class, 'index'])
     ->name('upload.form');

// Handle the file upload
Route::post('/upload', [DocumentController::class, 'upload'])
     ->name('upload.submit');

// Result page
Route::get('/result', [DocumentController::class, 'result'])
     ->name('upload.result');

// Download the ZIP
Route::get('/download', [DocumentController::class, 'download'])
     ->name('document.download');