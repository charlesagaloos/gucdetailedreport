<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Shopee Platform</title>
    <!-- DataTables CSS and JS -->
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/jquery.dataTables.css">
    <script type="text/javascript" charset="utf8" src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script type="text/javascript" charset="utf8" src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.js"></script>
    <!-- DataTables Date Range Plugin -->
    <script type="text/javascript" src="https://cdn.datatables.net/plug-ins/1.10.25/sorting/datetime-moment.js"></script>
    <script type="text/javascript" src="https://cdn.datatables.net/plug-ins/1.10.25/filtering/datetime-moment.js"></script>
    <!-- SheetJS (XLSX) library -->
    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.4/xlsx.full.min.js"></script>
    <!-- Moment.js library for date parsing/formatting -->
    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.1/moment.min.js"></script>
</head>
<body>
<button onclick="goBack()">Go Back</button>
    <h1>Tiktok</h1>

    <!-- Button to import Excel file -->
    <input type="file" id="fileInput" accept=".xls, .xlsx">
    <button onclick="importExcel()">Import Excel</button>

    <!-- Add Date From and Date To input fields -->
    <label for="dateFrom">Date From:</label>
    <input type="date" id="dateFrom">

    <label for="dateTo">Date To:</label>
    <input type="date" id="dateTo">

    <button onclick="filterData()">Filter Data</button>


    <!-- DataTable to display data -->
    <table id="dataTable">
        <thead>
            <tr>
                <th>Order ID</th>
                <th>Order Status</th>
                <th>Cancel Reason</th>
                <th>Return Refund</th>
                <th>Failed Deliver Status</th>
                <th>Tracking Number*</th>
                <th>Shipping Option</th>
                <th>Shipment Method</th>
                <th>Estimated Ship Out Date</th>
                <th>Ship Time</th>
                <th>Order Creation Date</th>
                <th>Order Paid Time</th>
                <th>Parent SKU Reference No.</th>
                <th>Product Name</th>
                <!-- Add more columns as needed -->
            </tr>
        </thead>
        <tbody>
            <!-- Data will be populated dynamically -->
        </tbody>
    </table>

    <script>
        // Function to go back to index.php
        function goBack() {
            // Assuming you want to go back to index.php
            window.location.href = 'index.php';
        }

        // Define a global variable to store the imported data
        var importedData = [];

        // Function to handle Excel file import
        function importExcel() {
            // ... (unchanged code)

            // Store the imported data globally
            importedData = jsonData;

            // Display data in DataTable
            displayDataInTable(importedData);
        }

        // Function to filter data based on date range
        function filterData() {
            // Get Date From and Date To values
            var dateFrom = document.getElementById('dateFrom').value;
            var dateTo = document.getElementById('dateTo').value;

            // Filter data based on the date range
            var filteredData = importedData.filter(function (row) {
                var orderDate = moment(row[10], 'YYYY/MM/DD'); // Assuming Order Creation Date is in the 11th column
                return orderDate.isSameOrAfter(dateFrom) && orderDate.isSameOrBefore(dateTo);
            });

            // Display filtered data in DataTable
            displayDataInTable(filteredData);
        }

        // Function to handle Excel file import
        function importExcel() {
            var fileInput = document.getElementById('fileInput');
            
            if (fileInput.files.length > 0) {
                var file = fileInput.files[0];
                var reader = new FileReader();

                reader.onload = function (e) {
                    var data = new Uint8Array(e.target.result);
                    var workbook = XLSX.read(data, { type: 'array' });

                    // Assuming the first sheet is the one with data
                    var sheetName = workbook.SheetNames[0];
                    var sheet = workbook.Sheets[sheetName];

                    // Convert sheet data to JSON starting from the second row
                    var jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, range: 1 });

                    // Display data in DataTable
                    displayDataInTable(jsonData);
                };

                reader.readAsArrayBuffer(file);
            } else {
                alert('Please select an Excel file.');
            }
        }

        function displayDataInTable(data) {
            var dataTable = $('#dataTable').DataTable();
            dataTable.clear().draw();

            for (var i = 0; i < data.length; i++) {
                dataTable.row.add(data[i]).draw(false);
            }
        }


        // Function to display data in DataTable
        function displayDataInTable(data) {
            var dataTable = $('#dataTable').DataTable();
            dataTable.clear().draw();

            for (var i = 0; i < data.length; i++) {
                dataTable.row.add(data[i]).draw(false);
            }
        }

        // DataTable initialization
        $(document).ready(function() {
            $('#dataTable').DataTable();
        });
    </script>

</body>
</html>
