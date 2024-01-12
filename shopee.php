<?php require 'config.php'; ?>
<!DOCTYPE html>
<html lang="en">
<head>
    <style>
        .even-row {
            background-color: #f2f2f2;
        }

        .odd-row {
            background-color: #ffffff;
        }

        .container2 {
            padding-top:20px;
            width: 90%; /* Adjust the width as needed */
            margin: auto; /* Center the container */
        }

        #dataTable {
            font-size: 12px;
            width: 100%;
        }

        #dataTable_wrapper {
            max-width: 100%;
        }

        .dt-buttons {
            margin-bottom: 10px;
        }
        .navbar-custom {
            background-color: #ee4d2d; /* Shopee orange color */
        }

        .navbar-brand {
            color: #ffffff !important;
        }
        #goBackLink {
            font-size: 14px;
            margin-right: 1rem;
            color: #ffffff; /* Blue color for the link */
        }
        /* Custom style for "Choose File" button */
        .custom-file-container {
            position: relative;
            overflow: hidden;
            display: inline-block;
        }

        .custom-file-input {
            position: absolute;
            font-size: 100px;
            right: 0;
            top: 0;
            opacity: 0;
            cursor: pointer;
            width:50%;
        }

        .custom-file-label {
            display: inline-block;
            padding: 8px 12px; /* Adjust padding as needed */
            font-size: 16px;
            font-weight: bold;
            cursor: pointer;
            background-color: #007bff;
            color: #fff;
            border: 1px solid #007bff;
            border-radius: 4px;
            max-width: 100%; /* Set the maximum width as needed */
            overflow: hidden;
            white-space: nowrap;
            text-overflow: ellipsis;
            
        }


        .custom-file-label:hover {
            background-color: #0056b3;
            border-color: #0056b3;
        }

        .custom-file-label:active {
            background-color: #0056b3;
            border-color: #0056b3;
        }

        .custom-file-label:focus {
            outline: none;
            box-shadow: 0 0 0 0.2rem rgba(0, 123, 255, 0.25);
        }

        .mr-2{
            margin-right: 0.5rem!important;
            width:20%;
        }

    </style>

<meta charset="UTF-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Shopee Platform</title>
<!-- Add Bootstrap CSS -->
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css">
<!-- SheetJS (XLSX) library -->
<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.4/xlsx.full.min.js"></script>
<!-- Moment.js library for date parsing/formatting -->
<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.1/moment.min.js"></script>
<!-- DataTables CSS and JS -->
<link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/jquery.dataTables.css">
<script type="text/javascript" charset="utf8" src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script type="text/javascript" charset="utf8" src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.js"></script>
<!-- DataTables Date Range Filtering CSS -->
<link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/datetime/1.11.5/css/dataTables.dateTime.min.css">
<!-- DataTables Date Range Filtering JS -->
<script type="text/javascript" src="https://cdn.datatables.net/datetime/1.11.5/js/dataTables.dateTime.min.js"></script>
<!-- Flatpickr CSS and JS -->
<link rel="stylesheet" type="text/css" href="https://cdn.jsdelivr.net/npm/flatpickr/dist/flatpickr.min.css">
<script type="text/javascript" src="https://cdn.jsdelivr.net/npm/flatpickr"></script>
<!-- SweetAlert library -->
<script src="https://cdn.jsdelivr.net/npm/sweetalert2@10"></script>

</head>
<body>

    <header>
        <nav class="navbar navbar-custom">
        <a id="goBackLink" href="index.php">Go Back</a>
            <span class="navbar-brand mb-0 h1">Shopee Platform</span>
        </nav>
    </header>

    <div class="container2">
        <h1>ORDERS</h1>
        <br>
        <div class="row mb-3">
            <div class="col">
                <div class="d-flex align-items-start">
                    <div class="custom-file-container mr-2">
                    <label id="fileInputLabel" for="fileInput" class="custom-file-label">
                            Choose File
                        </label>
                        <input type="file" id="fileInput" accept=".xls, .xlsx" class="custom-file-input" onchange="updateFilename()">
                    </div>
                    <button onclick="importExcel()" class="btn btn-primary">Import Excel</button>
                    <form method="post" action="shopee_guc_excel.php" onsubmit="updateDateValues()">
                    <!-- Add hidden inputs for Date From and Date To -->
                    <input type="hidden" id="hiddenDateFrom" name="dateFrom" value="">
                    <input type="hidden" id="hiddenDateTo" name="dateTo" value="">
                    <input type="hidden" id="hiddenOrderStatus" name="orderStatus" value="">
                    <input type="hidden" id="hiddenOrderStatusFilter" name="OrderStatusFilter" value="">
                    <button type="submit" name="generate_excel" class="btn btn-success ml-2">Generate Excel</button>
                </form>
                </div>
            </div>
        </div>
    </div>
    <div class="container2" style="width:90%; center">
        <!-- Date filter inputs -->
        <div class="row mt-3">
                <div class="col">
                    <label>Date From:</label>
                    <input type="text" id="dateFrom" class="form-control">
                </div>
                <div class="col">
                    <label>Date To:</label>
                    <input type="text" id="dateTo" class="form-control">
                </div>
                <div class="col">
                    <label for="orderStatusFilter">Order Status:</label>
                        <select class="form-control" id="orderStatusFilter">
                            <option value="">All</option>
                            <option value="Completed">Completed</option>
                            <option value="Cancelled">Cancelled</option>
                            <option value="Refund">Refund</option>
                            <option value="Shipping">Shipping</option>
                        </select>
                </div>
            </div>
            <br>
            <!-- DataTable to display data -->
            <table id="dataTable" class="table table-striped">
                <thead>
                    <tr>
                        <td>Order ID</td>
                        <td>Order Paid Date</td>
                        <td>Order Complete Date</td>
                        <td>Order Status</td>
                        <td>Tracking Number</td>
                        <td>Product Name</td>
                    </tr>
                </thead>
                <tbody>
                    <?php
                        $i = 1;
                        $rows = mysqli_query($conn, "SELECT Order_ID, MIN(Order_Creation_Date) as Order_Creation_Date, MIN(Order_Complete_Time) as Order_Complete_Time, MIN(Order_Status) as Order_Status, MIN(Cancel_Reason) as Cancel_Reason, MIN(Return_Refund_Status) as Return_Refund_Status, MIN(Tracking_Number) as Tracking_Number, MIN(Product_Name) as Product_Name FROM tbl_shopee GROUP BY Order_ID");

                        foreach ($rows as $row) :
                            ?>
                            <tr>
                            <td><a href="order_details.php?orderId=<?php echo $row["Order_ID"]; ?>"><?php echo $row["Order_ID"]; ?></a></td>
                            <td><?php echo $row["Order_Creation_Date"]; ?></td>
                            <td><?php echo $row["Order_Complete_Time"]; ?></td>
                            <td style="color: <?php echo ($row["Order_Status"] === 'Completed') ? 'green' : (($row["Order_Status"] === 'Refund') ? 'blue' : (($row["Order_Status"] === 'Shipping') ? 'orange' : 'red')); ?>">

                                <?php echo $row["Order_Status"]; ?>
                            </td>
                            <td><?php echo $row["Tracking_Number"]; ?></td>
                            <td><?php echo $row["Product_Name"]; ?></td>
                            </tr>
                    <?php endforeach; ?>
                </tbody>
            </table>
    </div>
    <footer class="text-center mt-5">
        <p>&copy; 2023 Charles. All rights reserved.</p>
    </footer>
    
    <script>

        function updateDateValues() {
                // Get the selected date values
                var dateFromValue = document.getElementById('dateFrom').value;
                var dateToValue = document.getElementById('dateTo').value;

                // Update the hidden input fields with the selected date values
                document.getElementById('hiddenDateFrom').value = dateFromValue;
                document.getElementById('hiddenDateTo').value = dateToValue;
        }

        // Function to go back to index.php
        function goBack() {
            // Assuming you want to go back to index.php
            window.location.href = 'index.php';
        }

        // Function to update the filename in the label
        function updateFilename() {
            var fileInput = document.getElementById('fileInput');
            var fileInputLabel = document.getElementById('fileInputLabel');

            // Check if a file is selected
            if (fileInput.files.length > 0) {
                // Update label with the selected filename
                fileInputLabel.innerText = fileInput.files[0].name;
            } else {
                // Reset label text if no file is selected
                fileInputLabel.innerText = 'Choose File';
            }
        }

        // Function to handle Excel file import
        function importExcel() {
            var fileInput = document.getElementById('fileInput');
            var file = fileInput.files[0];

            var reader = new FileReader();
            reader.onload = function(e) {
                var data = new Uint8Array(e.target.result);
                var workbook = XLSX.read(data, { type: 'array' });
                var sheetName = workbook.SheetNames[0];
                var sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, range: 1 });


                // Send the data to your server for processing
                $.ajax({
                    type: 'POST',
                    url: 'shopee_import.php',
                    data: JSON.stringify(sheetData),
                    success: function(response) {
                        alert('Data imported successfully!');
                        // Reload or update the DataTable with the new data
                        // You may need to modify this based on how your DataTable is initialized
                        location.reload();
                    },
                    contentType: 'application/json',
                    dataType: 'json'
                });
            };

            reader.readAsArrayBuffer(file);
        }

        // Function to display data in DataTable
        function displayDataInTable(data) {
            var dataTable = $('#dataTable').DataTable();
            dataTable.clear().draw();

            // Exclude the first row (index 0) which contains column names
            for (var i = 1; i < data.length; i++) {
                // Add a link to the Order ID column
                data[i][0] = '<a href="order_details.php?orderId=' + data[i][0] + '">' + data[i][0] + '</a>';

                dataTable.row.add(data[i]).draw(false);
            }
        }

        // DataTable initialization
        $(document).ready(function() {

            // Initialize Flatpickr for date inputs
            flatpickr("#dateFrom", {
                dateFormat: "F j, Y",
                onChange: function(selectedDates, dateStr, instance) {
                    dataTable.draw();
                }
            });

            flatpickr("#dateTo", {
                dateFormat: "F j, Y",
                onChange: function(selectedDates, dateStr, instance) {
                    dataTable.draw();
                }
            });

            // Add event listener to the DataTable search event
            $('#dataTable').on('search.dt', function () {
                var searchValue = $('#dataTable_filter input').val();
                $('#hiddenOrderStatus').val(searchValue);
            });

            // Add event listener to the Order Status filter
            $('#orderStatusFilter').on('change', function() {
                var selectedOrderStatus = $(this).val();

                // Use DataTable's search method to filter data based on Order Status
                dataTable.search(selectedOrderStatus).draw();
                $('#hiddenOrderStatusFilter').val(selectedOrderStatus);
            });

            var dataTable = $('#dataTable').DataTable({
                // Specify date format for columns containing dates
                columnDefs: [
                    {
                        targets: [1, 2],
                        render: function(data, type, row) {
                            // Assuming the date is in the format 'YYYY-MM-DD HH:mm'
                            var momentDate = moment(data, 'YYYY-MM-DD HH:mm');

                            // Check if the date is valid
                            var isValidDate = momentDate.isValid();

                            // Return the formatted date if valid, otherwise return an empty string
                            return isValidDate ? momentDate.format('MMMM D, YYYY') : '';
                        }
                    }
                ],
            });

            // Add Flatpickr date range filtering
            $.fn.dataTable.ext.search.push(
                function(settings, data, dataIndex) {
                    var minDate = $("#dateFrom").val();
                    var maxDate = $("#dateTo").val();
                    var orderDate = moment(data[1], 'MMMM D, YYYY'); // Assuming your date column is at index 1

                    if ((minDate === '' || orderDate.isSameOrAfter(minDate)) && (maxDate === '' || orderDate.isSameOrBefore(maxDate))) {
                        return true;
                    }

                    return false;
                }
            );

            // Add event listener to the form for "Generate Excel" button
            $('form').submit(function (event) {
                var dateFrom = $('#dateFrom').val();
                var dateTo = $('#dateTo').val();

                if (dateFrom === '' || dateTo === '') {
                    // Show SweetAlert warning if date filter is not set
                    Swal.fire({
                        icon: 'warning',
                        title: 'Date Filter Required',
                        text: 'Please set a date range before generating the Excel file.',
                    });

                    // Prevent the form submission
                    event.preventDefault();
                } else {
                    // Show SweetAlert for file downloading
                    Swal.fire({
                        title: 'Downloading',
                        text: 'Please wait...',
                        allowOutsideClick: false,
                        showConfirmButton: false,
                        onBeforeOpen: () => {
                            Swal.showLoading();
                        },
                    });

                    $.ajax({
                        type: 'POST',
                        url: 'shopee_guc_excel.php', // Update with your actual URL
                        data: $('form').serialize(),
                        success: function(response) {
                            // Handle success, if needed
                        },
                        complete: function(response) {
                            // Hide SweetAlert on completion
                            Swal.close();
                        }
                    });
                }
            });
        });
    </script>

    <!-- Add Bootstrap JS -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js"></script>

</body>
</html>
