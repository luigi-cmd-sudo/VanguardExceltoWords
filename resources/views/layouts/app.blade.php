<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>Excel to Word Converter</title>
    <link rel="stylesheet" href="{{ asset('css/app.css') }}">
    <link rel="stylesheet" href="{{ asset('css/header.css') }}">
</head>
<body>

    @include('partials.header')

    @yield('content')

    <script>
    document.querySelector('form').addEventListener('submit', function () 
        var btn = document.getElementById('submitBtn');
        btn.textContent = '⏳ Processing, please wait...';
        btn.disabled = true;

        function toggleMenu() {
        const nav = document.querySelector('.nav-menu');
        nav.classList.toggle('active');
    });

    
</script>
    
</body>
</html>