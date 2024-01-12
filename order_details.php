<?php
// order_details.php

require 'config.php';

if (isset($_GET['orderId'])) {
    $orderId = $_GET['orderId'];
    $orderStatusQuery = mysqli_query($conn, "SELECT Order_Status FROM tbl_shopee WHERE Order_ID = '$orderId'");
    
    if ($orderStatusQuery) {
        $orderStatusRow = mysqli_fetch_assoc($orderStatusQuery);
        $orderStatus = $orderStatusRow['Order_Status'];
    }

} else {
    // Handle the case when Order ID is not provided
    echo "Order ID not provided.";
}
?>
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
        .containeradjust {
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
        .subtotal-label-container {
        text-align: right;
        margin-top: 10px; /* Adjust the margin as needed */
    }

    .subtotal-label {
        font-weight: bold;
    }
    </style>

    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Shopee Platform</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/flatpickr/dist/flatpickr.min.css">
    <!-- Add Bootstrap CSS -->
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css">
    <!-- SheetJS (XLSX) library -->
    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.4/xlsx.full.min.js"></script>
    <!-- Moment.js library for date parsing/formatting -->
    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.1/moment.min.js"></script>
    <!-- DataTables CSS and JS -->
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.11.5/css/jquery.dataTables.css">
    <!-- Include DataTables JS -->
    <script type="text/javascript" charset="utf8" src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script type="text/javascript" charset="utf8" src="https://cdn.datatables.net/1.11.5/js/jquery.dataTables.js"></script>
    <!-- DataTables Date Range Filtering CSS -->
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/datetime/1.11.5/css/dataTables.dateTime.min.css">
    <!-- DataTables Date Range Filtering JS -->
    <script type="text/javascript" src="https://cdn.datatables.net/datetime/1.11.5/js/dataTables.dateTime.min.js"></script>
    <!-- Add Swal (SweetAlert) JS -->
<script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>
</head>
<body>

    <header>
        <nav class="navbar navbar-custom">
        <a id="goBackLink" href="shopee.php">Go Back</a>
            <span class="navbar-brand mb-0 h1">Shopee Platform</span>
        </nav>
    </header>

    <div class="container2">
        <h2>ORDER DETAILS</h2>
    </div>
    <div class="container2" style="width:90%; center">
                <!-- DataTable to display data -->
                <table id="dataTable" class="table table-striped">
            <thead>
                <tr>
                    <td>Order ID</td>
                    <td>Order Paid Date</td>
                    <td>Order Complete Date</td>
                    <td>Order Status</td>
                    <?php if (strtolower($orderStatus) === "cancelled") : ?>
                        <td>Cancel Reason</td>
                        <td>Failed Deliver Status</td>
                        <td>Return / Refund Status</td>
                    <?php else : ?>
                        <td></td> <!-- Empty cell if the order is not cancelled -->
                    <?php endif; ?>
                    <td>Tracking Number</td>
                    <td>Product Name</td>
                    <td>SKU Reference No.</td>
                </tr>
            </thead>
            <tbody>
                <?php
                    $i = 1;
                    $rows = mysqli_query($conn, "SELECT * FROM tbl_shopee WHERE Order_ID = '$orderId'");
                    foreach ($rows as $row) :
                        ?>
                        <tr>
                        <td><?php echo $row["Order_ID"]; ?></td>
                        <td><?php echo $row["Order_Paid_Time"]; ?></td>
                        <td><?php echo $row["Order_Complete_Time"]; ?></td>
                        <td><?php echo $row["Order_Status"]; ?></td>
                        <?php if (strtolower($orderStatus) === "cancelled") : ?>
                            <td><?php echo $row["Cancel_Reason"]; ?></td>
                            <td><?php echo $row["Return_Refund_Status"]; ?></td>
                            <td><?php echo "" ?></td>
                        <?php else : ?>
                            <td></td> <!-- Empty cell if the order is not cancelled -->
                        <?php endif; ?>
                        <td><?php echo $row["Tracking_Number"]; ?></td>
                        <td><?php echo $row["Product_Name"]; ?></td>
                        <td><?php echo $row["SKR_No"]; ?></td>
                        </tr>
                <?php endforeach; ?>
            </tbody>
        </table>
    </div>

    <div class="container2">
        <h2>PRODUCT PRICE</h2>
    </div>
    <div class="container2" style="width:90%; center">
                <!-- productPricesTable to display data -->
                <table id="productPricesTable" class="table table-striped">
            <thead>
                <tr>
                    <td>Original Price</td>
                    <td>Deal Price</td>
                    <td>Quantity</td>
                    <td>Total Discount</td>
                    <td>Product Subtotal</td>
                </tr>
            </thead>
            <tbody>
                <?php
                    $i = 1;
                    $rows = mysqli_query($conn, "SELECT * FROM tbl_shopee WHERE Order_ID = '$orderId'");
                    foreach ($rows as $row) :
                        ?>
                        <tr>
                        <td><?php echo number_format($row["Orig_Price"], 2); ?></td>
                        <td><?php echo number_format($row["Deal_Price"], 2); ?></td>
                        <td><?php echo $row["Qnty"]; ?></td>
                        <td><?php echo number_format($row["Total_Discount"], 2); ?></td>
                        <td><?php echo number_format($row["Product_Subtotal"], 2); ?></td>

                        </tr>
                <?php endforeach; ?>
            </tbody>
        </table>
                <!-- Subtotal label outside the table -->
                <?php
                $sqlGet = "SELECT SUM(Product_Subtotal) AS totalSubtotal FROM tbl_shopee WHERE Order_ID = '$orderId'";
                $resultGet = mysqli_query($conn, $sqlGet);

                if ($resultGet) {
                    $row = mysqli_fetch_assoc($resultGet);
                    $productSubtotal = $row['totalSubtotal'];
                } else {
                    // Handle the case where the query fails
                    $productSubtotal = 0; // Set a default value or handle the error as needed
                }
            ?>
            <div class="subtotal-label-container">
                <label class="subtotal-label"><h2>Subtotal: <?php echo number_format($productSubtotal, 2); ?></h2></label>
            </div>
    </div>

    <div class="container2">
        <h2>DEDUCTIONS</h2>
    </div>
    <div class="container2" style="width:90%; center">
                <!-- deductionsTable to display data -->
                <table id="deductionsTable" class="table table-striped">
            <thead>
                <tr>
                    <td>Shipping Fee Discount</td>
                    <td>Seller Voucher</td>
                    <td>Commission Fee</td>
                    <td>Service Fee</td>
                    <td>Transaction Fee</td>
                    <td>Total Deductions</td>
                    <td>Net Sales</td>
                </tr>
            </thead>
            <tbody>
                <?php
                    $i = 1;
                    $rows = mysqli_query($conn, "SELECT Order_ID, D_Shipping_Fee_Discount, Seller_Voucher, D_Comission_Fee_Discount, Service_Fee , D_Transaction_Fee, D_Total_Deductions, Net_Sales FROM tbl_shopee WHERE Order_ID = '$orderId' LIMIT 1");
                    foreach ($rows as $row) :
                        ?>
                        <tr class="edit-deductions" data-order-id="<?php echo $row["Order_ID"]; ?>">
                        <td><?php echo isset($row["D_Shipping_Fee_Discount"]) ? number_format($row["D_Shipping_Fee_Discount"], 2) : 0; ?></td>
                        <td><?php echo isset($row["Seller_Voucher"]) ? number_format($row["Seller_Voucher"], 2) : 0; ?></td>
                        <td><?php echo isset($row["D_Comission_Fee_Discount"]) ? number_format($row["D_Comission_Fee_Discount"], 2) : 0; ?></td>
                        <td><?php echo isset($row["Service_Fee"]) ? number_format($row["Service_Fee"], 2) : 0; ?></td>
                        <td><?php echo isset($row["D_Transaction_Fee"]) ? number_format($row["D_Transaction_Fee"], 2) : 0; ?></td>
                        <td><?php echo isset($row["D_Total_Deductions"]) ? number_format($row["D_Total_Deductions"], 2) : 0; ?></td>
                        <td><?php echo isset($row["Net_Sales"]) ? number_format($row["Net_Sales"], 2) : 0; ?></td>
                        </tr>
                <?php endforeach; ?>
            </tbody>
        </table>
    </div>
    
    <div class="containerAdjust" id="adjustmentTableContainer">
        <div class="container2" style="width:90%; center" >
            <h2>ADJUSTMENTS</h2>
            <!-- adjustmentTable to display data -->
            <table id="adjustmentTable" class="table table-striped">
                <thead>
                    <tr>
                        <td>Refund</td>
                        <td>Adjustment Complete Date</td>
                        <td>Total Adjustment Amount</td>
                        <td>Return</td>
                        <td>Reimbursed</td>
                        <td>Date Submitted</td>
                        <td>Status</td>
                    </tr>
                </thead>
                <tbody>
                    <?php
                        $i = 1;
                        $rows = mysqli_query($conn, "SELECT * FROM tbl_shopee WHERE Order_ID = '$orderId' LIMIT 1");
                        foreach ($rows as $row) :
                            ?>
                            <tr class="edit-adjustments" data-order-id="<?php echo $row["Order_ID"]; ?>">
                                <td><?php echo isset($row["R_Refund_Amount"]) ? number_format($row["R_Refund_Amount"], 2) : 0; ?></td>
                                <td><?php echo !empty($row["A_Adjustment_Complete_Date"]) ? $row["A_Adjustment_Complete_Date"] : "N/A"; ?></td>
                                <td><?php echo isset($row["A_Total_Adjustment_Amount"]) ? number_format($row["A_Total_Adjustment_Amount"], 2) : 0; ?></td>
                                <td><?php echo !empty($row["A_Return"]) ? $row["A_Return"] : "N/A"; ?></td>
                                <td><?php echo !empty($row["A_Reimbursed"]) ? $row["A_Reimbursed"] : "N/A"; ?></td>
                                <td><?php echo !empty($row["A_Date_Submitted"]) ? $row["A_Date_Submitted"] : "N/A"; ?></td>
                                <td><?php echo !empty($row["A_Status"]) ? $row["A_Status"] : "N/A"; ?></td>
                            </tr>
                    <?php endforeach; ?>
                </tbody>
            </table>
        </div>
    </div>
    
    <!-- Deductions Modal -->
    <div class="modal" id="deductionsModal">
        <div class="modal-dialog">
            <div class="modal-content">
                <!-- Modal Header -->
                <div class="modal-header">
                    <h4 class="modal-title">Deductions</h4>
                    <button type="button" class="close" data-dismiss="modal">&times;</button>
                </div>
                <form id="deductionsForm">
                    <!-- Modal Body -->
                    <div class="modal-body">
                        <form>
                            <div class="form-group">
                                <label for="shippingFeeDiscount">Shipping Fee Discount:</label>
                                <input type="text" class="form-control" id="shippingFeeDiscount" name="shippingFeeDiscount" placeholder="Enter Shipping Fee Discount" oninput="validateNumericInput(this)">
                            </div>
                            <div class="form-group">
                                <label for="sellerVoucher">Seller Voucher:</label>
                                <input type="text" class="form-control" id="sellerVoucher" name="sellerVoucher" placeholder="Enter Seller Voucher" oninput="validateNumericInput(this)">
                            </div>
                            <div class="form-group">
                                <label for="commissionFee">Commission Fee:</label>
                                <input type="text" class="form-control" id="commissionFee" name="commissionFee" placeholder="Enter Commission Fee" oninput="validateNumericInput(this)">
                            </div>
                            <div class="form-group">
                                <label for="serviceFee">Service Fee:</label>
                                <input type="text" class="form-control" id="serviceFee" name="serviceFee" placeholder="Enter Service Fee" oninput="validateNumericInput(this)">
                            </div>
                            <div class="form-group">
                                <label for="transactionFee">Transaction Fee:</label>
                                <input type="text" class="form-control" id="transactionFee" name="transactionFee" placeholder="Enter Transaction Fee" oninput="validateNumericInput(this)">
                            </div>
                            <div class="form-group">
                                <!-- <label for="totalDeductions">Total Deductions:</label> -->
                                <input type="hidden" class="form-control" id="totalDeductions" name="totalDeductions" placeholder="Enter Total Deductions" oninput="validateNumericInput(this)">
                            </div>
                            <div class="form-group">
                                <!-- <label for="netSales">Net Sales:</label> -->
                                <input type="hidden" class="form-control" id="netSales" name="netSales" placeholder="Enter Net Sales" oninput="validateNumericInput(this)">
                            </div>
                        </form>
                    </div>
                </form>
                
                <!-- Modal Footer -->
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
                    <!-- Add a button to save changes -->
                    <button type="button" class="btn btn-primary" onclick="saveDeductions()">Save Changes</button>
                </div>
            </div>
        </div>
    </div>

    <!-- Adjustments Modal -->
    <div class="modal" id="adjustmentsModal">
        <div class="modal-dialog">
            <div class="modal-content">
                <!-- Modal Header -->
                <div class="modal-header">
                    <h4 class="modal-title">Adjustments</h4>
                    <button type="button" class="close" data-dismiss="modal">&times;</button>
                </div>
                <form id="adjustmentForm">
                    <!-- Modal Body -->
                    <div class="modal-body">
                        <form>
                            <div class="form-group">
                                <label for="adjustmentRefund">Refund:</label>
                                <input type="text" class="form-control" id="adjustmentRefund" name="adjustmentRefund" placeholder="Refund" oninput="validateNumericInput(this)">
                            </div>

                            <div class="form-group">
                                <label for="adjustmentCompleteDate">Adjustment Complete Date:</label>
                                <input type="text" class="form-control" id="adjustmentCompleteDate" name="adjustmentCompleteDate">
                            </div>
                            <div class="form-group">
                                <label for="totalAdjustmentAmount">Total Adjustment Amount:</label>
                                <input type="text" class="form-control" id="totalAdjustmentAmount" name="totalAdjustmentAmount" placeholder="Total Adjustment Amount" oninput="validateNumericInput(this)">
                            </div>
                            <div class="form-group">
                                <label for="adjusmentReturn">Return:</label>
                                <select class="form-control" id="adjusmentReturn" name="adjusmentReturn">
                                    <option value="" selected>- Select -</option>
                                    <option value="Returned">Returned</option>
                                </select>
                            </div>
                            <div class="form-group">
                                <label for="adjustmentReimbursed">Reimbursed:</label>
                                <select class="form-control" id="adjustmentReimbursed" name="adjustmentReimbursed">
                                    <option value="" selected>- Select -</option>
                                    <option value="CREDITED">CREDITED</option>
                                    <option value="WAITING FOR SHOPEE'S AGENT">WAITING FOR SHOPEE'S AGENT</option>
                                </select>
                            </div>
                            <div class="form-group">
                                <label for="adjustmentDateSubmitted">Date Submitted:</label>
                                <input type="text" class="form-control" id="adjustmentDateSubmitted" name="adjustmentDateSubmitted">
                            </div>
                            <div class="form-group">
                                <label for="adjustmentStatus">Status:</label>
                                <select class="form-control" id="adjustmentStatus" name="adjustmentStatus">
                                    <option value="" selected >- Select -</option>
                                    <option value="OVERDUE">OVERDUE</option>
                                    <option value="REMITTED">REMITTED</option>
                                    <option value="SUBMITTED">SUBMITTED</option>
                                    <option value="DONE">DONE</option>
                                </select>
                            </div>
                        </form>
                    </div>
                </form>
                
                <!-- Modal Footer -->
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
                    <!-- Add a button to save changes -->
                    <button type="button" class="btn btn-primary" onclick="saveAdjustment()">Save Changes</button>
                </div>
            </div>
        </div>
    </div>
    </div>
    <script src="https://cdn.jsdelivr.net/npm/flatpickr"></script>
    <script>
        flatpickr("#adjustmentDateSubmitted", {
            dateFormat: "F j, Y",
            onChange: function(selectedDates, dateStr, instance) {
                // You can perform additional actions when the date changes
                // console.log(dateStr);
            }
        });
        flatpickr("#adjustmentCompleteDate", {
            dateFormat: "F j, Y",
            onChange: function(selectedDates, dateStr, instance) {
                // You can perform additional actions when the date changes
                // console.log(dateStr);
            }
        });
        $('#deductionsModal, #adjustmentsModal').on('show.bs.modal', function(event) {
            var orderId = $(this).data('order-id');
            $(this).find('#orderIdPlaceholder').text(orderId);
            // Disable the "Net Sales" input field
            $('#netSales').prop('readonly', true);
            $('#totalDeductions').prop('readonly', true);
        });
        // Function to go back to index.php
        function goBack() {
            // Assuming you want to go back to index.php
            window.location.href = 'index.php';
        }

        $(document).ready(function() {

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
            $('#deductionsTable').on('click', '.edit-deductions', function() {
                var orderId = $(this).data('order-id');
                var row = $(this).closest('tr');
                var shippingFeeDiscount = row.find('td:eq(0)').text();
                var sellerVoucher = row.find('td:eq(1)').text();
                var comm_fee = row.find('td:eq(2)').text();
                var service_fee = row.find('td:eq(3)').text();
                var transaction_fee = row.find('td:eq(4)').text();
                var total_deductions = row.find('td:eq(5)').text();
                var net_sales = row.find('td:eq(6)').text();

                // Set the modal field values
                $('#shippingFeeDiscount').val(shippingFeeDiscount);
                $('#sellerVoucher').val(sellerVoucher);
                $('#commissionFee').val(comm_fee);
                $('#serviceFee').val(service_fee);
                $('#transactionFee').val(transaction_fee);
                $('#totalDeductions').val(total_deductions);
                $('#netSales').val(net_sales);


                $('#deductionsModal').data('order-id', orderId).modal('show');
                // Add code to populate modal fields with data for editing
            });

            $('#adjustmentTable').on('click', '.edit-adjustments', function() {
                var orderId = $(this).data('order-id');
                var row = $(this).closest('tr');
                var a_Refund = row.find('td:eq(0)').text();
                var adjustedRefund = a_Refund.trim() !== "0.00" ? parseFloat(a_Refund) : "";
                var a_adjustmentCompleteDate = row.find('td:eq(1)').text();
                var adjustedCompleteDate = a_adjustmentCompleteDate !== "N/A" ? a_adjustmentCompleteDate : "";
                var a_totalAdjustmentAmount = row.find('td:eq(2)').text();
                var adjustedAmount = a_totalAdjustmentAmount.trim() !== "0.00" ? parseFloat(a_totalAdjustmentAmount) : "";
                var a_return = row.find('td:eq(3)').text();
                var adjustedValue = a_return.trim() !== "N/A" ? parseFloat(a_return) : "";
                var a_reimbursed = row.find('td:eq(4)').text();
                var adjustedValueReimbursed = a_reimbursed.trim() !== "N/A" ? parseFloat(a_reimbursed) : "";
                var a_dateSubmitted = row.find('td:eq(5)').text();
                var adjustedDateSubmitted = a_dateSubmitted !== "N/A" ? a_dateSubmitted : "";
                var a_Status = row.find('td:eq(6)').text();
                var adjustedStatus = a_Status.trim() !== "N/A" ? parseFloat(a_Status) : "";

                // Set the modal field values
                $('#adjustmentRefund').val(a_Refund);
                $('#adjustmentCompleteDate').val(adjustedCompleteDate);
                $('#totalAdjustmentAmount').val(adjustedAmount);
                $('#adjusmentReturn').val(adjustedValue);
                $('#adjustmentReimbursed').val(adjustedValueReimbursed);
                $('#adjustmentDateSubmitted').val(adjustedDateSubmitted);
                $('#adjustmentStatus').val(adjustedStatus);

                $('#adjustmentsModal').data('order-id', orderId).modal('show');
                // Add code to populate modal fields with data for editing
            });

            // var orderStatus = "<?php echo $orderStatus; ?>";

            // Check if the orderStatus is "cancelled" and toggle the adjustment table visibility
            // comment ko muna kasi need ko sya for now
            // if (orderStatus.toLowerCase() === "cancelled") {
            //     $('#adjustmentTableContainer').show();  // Show the adjustment table
            // } else {
            //     $('#adjustmentTableContainer').hide();  // Hide the adjustment table
            // }

            // DataTable initialization code (same as before)
            $('#dataTable').DataTable();
            $('#productPricesTable').DataTable();
            $('#deductionsTable').DataTable();
            $('#adjustmentTable').DataTable();

            
        });

        function validateNumericInput(input) {
            // Remove non-numeric characters
            input.value = input.value.replace(/[^0-9.]/g, '');

            // Ensure that the first character is not a dot
            if (input.value.charAt(0) === '.') {
                input.value = input.value.slice(1);
            }
        }
        
        function saveDeductions() {
            // Get the Order ID from the modal
            var orderId = $('#deductionsModal').data('order-id');

            // Prepare the data to be sent
            var formData = $('#deductionsForm').serialize();

            // Send the data to the server using AJAX
            $.ajax({
                type: 'POST',
                url: 'save_shopee_deductions.php',
                data: formData + '&orderId=' + orderId,
                dataType: 'json', // Specify the expected data type
                success: function (response) {
                    // Check if the response is valid JSON and has the expected properties
                    if (response && response.status && response.message) {
                        // Check the response status
                        if (response.status === 'success') {
                            // Show SweetAlert popup upon successful save
                            Swal.fire({
                                icon: 'success',
                                title: 'Success!',
                                text: response.message,
                            }).then(() => {
                                // Reload the entire webpage after successful save
                                location.reload();
                            });
                        } else {
                            // Show SweetAlert popup for errors
                            Swal.fire({
                                icon: 'error',
                                title: 'Error!',
                                text: response.message,
                            });
                        }
                    } else {
                        // Handle unexpected response format
                        console.error('Unexpected response format:', response);
                        Swal.fire({
                            icon: 'error',
                            title: 'Error!',
                            text: 'Unexpected response format.',
                        });
                    }
                },
                error: function (error) {
                    // Handle AJAX errors
                    console.error(error);
                    Swal.fire({
                        icon: 'error',
                        title: 'Error!',
                        text: 'An error occurred while processing your request.',
                    });
                },
            });
        }
        function saveAdjustment() {
            // Get the Order ID from the modal
            var orderId = $('#adjustmentsModal').data('order-id');

            // Prepare the data to be sent
            var formData = $('#adjustmentForm').serialize();

            // Send the data to the server using AJAX
            $.ajax({
                type: 'POST',
                url: 'save_shopee_adjustments.php',
                data: formData + '&orderId=' + orderId,
                dataType: 'json', // Specify the expected data type
                success: function (response) {
                    // Check if the response is valid JSON and has the expected properties
                    if (response && response.status && response.message) {
                        // Check the response status
                        if (response.status === 'success') {
                            // Show SweetAlert popup upon successful save
                            Swal.fire({
                                icon: 'success',
                                title: 'Success!',
                                text: response.message,
                            }).then(() => {
                                // Reload the entire webpage after successful save
                                location.reload();
                            });
                        } else {
                            // Show SweetAlert popup for errors
                            Swal.fire({
                                icon: 'error',
                                title: 'Error!',
                                text: response.message,
                            });
                        }
                    } else {
                        // Handle unexpected response format
                        console.error('Unexpected response format:', response);
                        Swal.fire({
                            icon: 'error',
                            title: 'Error!',
                            text: 'Unexpected response format.',
                        });
                    }
                },
                error: function (error) {
                    // Handle AJAX errors
                    console.error(error);
                    Swal.fire({
                        icon: 'error',
                        title: 'Error!',
                        text: 'An error occurred while processing your request.',
                    });
                },
            });
        }




            
    </script>

    <!-- Add Bootstrap JS -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js"></script>

</body>
</html>