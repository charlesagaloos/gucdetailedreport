<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SOZO</title>
    <!-- Favicon -->
    <link rel="icon" href="images/icon.jpg" type="image/x-icon">
    <style>
        body {
            font-family: Arial, sans-serif;
            background: linear-gradient(to bottom right, #ee4d2d, #0066c0, #7ab55c, #000000);
            margin: 0;
            padding: 0;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            height: 100vh;
        }

        h1 {
            color: white;
        }

        .shopee {
            background-color: #ee4d2d; 
        }

        .lazada {
            background-color: #0066c0; 
        }

        .shopify {
            background-color: #7ab55c;
        }

        .tiktok {
            background-color: #000000; 
        }

        .shopee, .lazada, .shopify, .tiktok {
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        }

        button {
            padding: 10px 20px;
            margin: 10px;
            font-size: 16px;
            cursor: pointer;
            color: #fff;
            border: none;
            border-radius: 5px;
            transition: background-color 0.3s, box-shadow 0.3s;
            flex: 1; /* Added this line to make buttons equally share the space */
        }

        button:hover {
            opacity: 0.8;
            box-shadow: 0 8px 16px rgba(0, 0, 0, 0.2);
        }
        h1 img {
            max-width: 30%;
            height: auto;
            display: block; 
            margin: 0 auto; 
        }
    </style>
</head>
<body>

    <div class="container">
        <h1><img src="images/logo.png" alt="GLOIRE UNLIMITED COMPANY"></h1>
    </div>

    <div style="display: flex; justify-content: space-around; width: 50%;">
        <button class="shopee" onclick="redirectToShopee()">Shopee</button>
        <button class="lazada" onclick="redirectToLazada()">Lazada</button>
        <button class="shopify" onclick="redirectToShopify()">Shopify</button>
        <button class="tiktok" onclick="redirectToTikTok()">TikTok</button>
    </div>

    <script>
        function redirectToShopee() {
            window.location.href = 'shopee.php';
        }

        function redirectToLazada() {
            window.location.href = 'lazada.php';
        }

        function redirectToShopify() {
            window.location.href = 'shopify.php';
        }

        function redirectToTikTok() {
            window.location.href = 'tiktok.php';
        }
    </script>

</body>
</html>
