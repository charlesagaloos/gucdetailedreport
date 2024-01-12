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

        // Extract the form data
        $adjustmentRefund = $_POST['adjustmentRefund'];
        $adjustmentCompleteDate = $_POST['adjustmentCompleteDate'];
        $totalAdjustmentAmount = isset($_POST['totalAdjustmentAmount']) ? floatval($_POST['totalAdjustmentAmount']) : 0;
        $adjusmentReturn = $_POST['adjusmentReturn'];
        $adjustmentReimbursed = $_POST['adjustmentReimbursed'];
        $adjustmentDateSubmitted = $_POST['adjustmentDateSubmitted'];
        $adjustmentStatus = $_POST['adjustmentStatus'];

        // Update the deductions in the database
        $updateQuery = "UPDATE tbl_shopee 
                        SET R_Refund_Amount = '$adjustmentRefund',
                            A_Adjustment_Complete_Date = '$adjustmentCompleteDate',
                            A_Total_Adjustment_Amount = '$totalAdjustmentAmount',
                            A_Return = '$adjusmentReturn',
                            A_Reimbursed = '$adjustmentReimbursed',
                            A_Date_Submitted = '$adjustmentDateSubmitted',
                            A_Status = '$adjustmentStatus'
                        WHERE Order_ID = '$orderId'";

        $result = mysqli_query($conn, $updateQuery);

        if ($result) {
            // If the update was successful, send a success response
            $response = array('status' => 'success', 'message' => 'Adjusments updated successfully.');
            echo json_encode($response);
        } else {
            // If the update failed, send an error response with details
            $error = mysqli_error($conn);
            $response = array('status' => 'error', 'message' => 'Failed to update Adjusments. ' . $error);
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
