@extends('layouts.app')

@section('title','Documents Ready')

@section('content')

<div class="card">

    <div class="icon">✅</div>
    <h1>Documents Ready!</h1>

    <p class="summary">
        <strong>{{ $rowCount }}</strong> Word document(s) have been generated
        and packaged into a single ZIP file.
    </p>

    {{--
        Download All Files button.
        Passes the zip path as a query parameter to the download route.
        urlencode() makes the path safe to put in a URL.
    --}}
    <a href="{{ route('document.download') }}?file={{ urlencode($zipPath) }}"
       class="btn btn-download">
        ⬇️ Download All Files (.zip)
    </a>

    {{--
        Submit Another button.
        Simply links back to the upload form — no JavaScript needed.
    --}}
    <a href="{{ route('upload.form') }}" class="btn btn-another">
        📂 Submit Another File
    </a>

    <p class="note">
        The ZIP file contains {{ $rowCount }} .docx file(s).
    </p>

</div>

@endSection

