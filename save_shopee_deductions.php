<?php
error_reporting(E_ALL);
ini_set('display_errors', 1);

// save_deductions.php

// Include the database connection and necessary functions
require 'config.php';

// Check if the request is a POST request
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    try {
        // Get the Order ID from the POST data
        $orderId = $_POST['orderId'];

       // Check the Order_Status
       $sqlStatus = "SELECT Order_Status FROM tbl_shopee WHERE Order_ID = '$orderId'";
       $resultStatus = mysqli_query($conn, $sqlStatus);

       if ($resultStatus) {
           $rowStatus = mysqli_fetch_assoc($resultStatus);
           $orderStatus = $rowStatus['Order_Status'];

           // Get the total Product_Subtotal for the specified Order_ID if Order_Status is not "Cancelled"
           if ($orderStatus !== 'Cancelled') {
               $sqlGet = "SELECT SUM(Product_Subtotal) AS totalSubtotal FROM tbl_shopee WHERE Order_ID = '$orderId'";
               $resultGet = mysqli_query($conn, $sqlGet);

               if ($resultGet) {
                   $row = mysqli_fetch_assoc($resultGet);
                   $productSubtotal = $row['totalSubtotal'];
               } else {
                   // Handle the case where the query fails
                   $productSubtotal = 0; // Set a default value or handle the error as needed
               }
           } else {
               // If Order_Status is "Cancelled," set $productSubtotal to 0
               $productSubtotal = 0;
           }
       } else {
           // Handle the case where the query for Order_Status fails
           $productSubtotal = 0; // Set a default value or handle the error as needed
       }

        // Extract the form data
        $shippingFeeDiscount = $_POST['shippingFeeDiscount'];
        $sellerVoucher = $_POST['sellerVoucher'];
        $commissionFee = $_POST['commissionFee'];
        $serviceFee = $_POST['serviceFee'];
        $transactionFee = $_POST['transactionFee'];

        // Calculate the total deductions
        $totalDeductions = $shippingFeeDiscount + $sellerVoucher + $commissionFee + $serviceFee + $transactionFee;
        
        // get the Net Sale
        $netSales = $productSubtotal-$totalDeductions;
        
        // Validate the data (perform additional validation as needed)

        // Update the deductions in the database
        $updateQuery = "UPDATE tbl_shopee 
                        SET D_Shipping_Fee_Discount = '$shippingFeeDiscount',
                            Seller_Voucher = '$sellerVoucher',
                            D_Comission_Fee_Discount = '$commissionFee',
                            Service_Fee = '$serviceFee',
                            D_Transaction_Fee = '$transactionFee',
                            D_Total_Deductions = '$totalDeductions',
                            Net_Sales = '$netSales'
                        WHERE Order_ID = '$orderId'";

        $result = mysqli_query($conn, $updateQuery);

        if ($result) {
            // If the update was successful, send a success response
            $response = array('status' => 'success', 'message' => 'Deductions updated successfully.');
            echo json_encode($response);
        } else {
            // If the update failed, send an error response with details
            $error = mysqli_error($conn);
            $response = array('status' => 'error', 'message' => 'Failed to update deductions. ' . $error);
            echo json_encode($response);
        }
    } catch (Exception $e) {
        // Handle any unexpected exceptions and send an error response
        $response = array('status' => 'error', 'message' => 'An unexpected error occurred.');
        echo json_encode($response);
    }
} else {
    // If the request is not a POST request, send an error response
    $response = array('status' => 'error', 'message' => 'Invalid request method.');
    echo json_encode($response);
}
?>
