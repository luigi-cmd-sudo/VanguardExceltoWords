@extends('layouts.app')

@section('title', 'Upload Excel File')

@section('content')

    <div class="card">

        <h1>📄 Excel → Word Converter</h1>
        <p class="subtitle">
            Upload an Excel file. One Word document will be generated per candidate row.
            All documents will be packaged into a single ZIP file for download.
        </p>

        @if ($errors->any())
            <div class="error-box">
                @foreach ($errors->all() as $error)
                    <p>⚠️ {{ $error }}</p>
                @endforeach
            </div>
        @endif

        <form method="POST" action="{{ route('upload.submit') }}" enctype="multipart/form-data">
            @csrf

            <label for="excel_file">Select Excel File</label>
            <input type="file" id="excel_file" name="excel_file" accept=".xlsx,.xls,.csv">
            <p class="hint">Accepted: .xlsx, .xls, .csv &nbsp;·&nbsp; Max size: 10 MB &nbsp;·&nbsp; Max rows: 200</p>

            <button type="submit" class="btn-submit" id="submitBtn">
                ⚡ Convert to Word Documents
            </button>
        </form>

        <div class="requirements">
            <strong>Excel file requirements:</strong>
            <ul>
                <li>Row 1 must be the header row</li>
                <li>Required columns: <code>Candidate Name</code>, <code>Case Reference Number</code>,
                    <code>Case Received Date</code>, <code>Report Date</code></li>
                <li>Dates can be Excel date format or text (e.g. 2026-04-14)</li>
                <li>Maximum <strong>200 data rows</strong> per upload</li>
                <li>Empty rows are skipped automatically</li>
            </ul>
        </div>

    </div>

@endSection

