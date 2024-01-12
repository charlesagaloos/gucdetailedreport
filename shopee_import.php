<?php
// error_reporting(E_ALL);
// ini_set('display_errors', 1);

ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);

// Database connection parameters
$servername = "localhost";
$username = "root";
$password = "";
$dbname = "db_gloire";

// Create connection
$conn = new mysqli($servername, $username, $password, $dbname);

// Check connection
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
}

// Get JSON data from the POST request
$jsonData = file_get_contents("php://input");

// Check if the input data is not empty
if (empty($jsonData)) {
    die("Error: No data received.");
}

$data = json_decode($jsonData, true);

// Check if JSON decoding was successful
if (json_last_error() !== JSON_ERROR_NONE) {
    die("Error decoding JSON data: " . json_last_error_msg() . "\nJSON Data: " . $jsonData);
}

// Check if $data is an array
if (!is_array($data)) {
    die("Invalid data format. Expected an array. JSON Data: " . $jsonData);
}

// Loop through the data and insert into the database
foreach ($data as $row) {
    $Order_ID = $conn->real_escape_string($row[0]);
    $Order_Status = $conn->real_escape_string($row[1]);
    $Cancel_Reason = $conn->real_escape_string($row[2]);
    $Return_Refund = $conn->real_escape_string($row[3]);
    $TrackingNo = $conn->real_escape_string($row[4]);
    $Ship_Option = $conn->real_escape_string($row[5]);
    $Ship_Method = $conn->real_escape_string($row[6]);
    $Estimated_ShipOut_Date = $conn->real_escape_string($row[7]);
    $ShipTime = $conn->real_escape_string($row[8]);
    $Order_Creation_Date = $conn->real_escape_string($row[9]);
    $Order_Paid_Time = $conn->real_escape_string($row[10]);
    $PSR_No = $conn->real_escape_string($row[11]);
    $Product_Name = $conn->real_escape_string($row[12]);
    $SKR_No = $conn->real_escape_string($row[13]);
    $Var_Name = $conn->real_escape_string($row[14]);
    $OrigPrice = $conn->real_escape_string($row[15]);
    $DealPrice = $conn->real_escape_string($row[16]);
    $Qnty = $conn->real_escape_string($row[17]);
    $Returned_Qnty = $conn->real_escape_string($row[18]);
    $Product_SubTotal = $conn->real_escape_string($row[19]);
    $Total_Disc = $conn->real_escape_string($row[20]);
    $Price_Disc_Seller = $conn->real_escape_string($row[21]);
    $Shopee_Rebate = $conn->real_escape_string($row[22]);
    $SKU_Total_Rebate = $conn->real_escape_string($row[23]);
    $No_Items_InOrder = $conn->real_escape_string($row[24]);
    $Order_Total_Weight = $conn->real_escape_string($row[25]);
    $Seller_Voucher = $conn->real_escape_string($row[26]);
    $SAC_Cashback = $conn->real_escape_string($row[27]);
    $Shopee_Voucher = $conn->real_escape_string($row[28]);
    $BundleDeals = $conn->real_escape_string($row[29]);
    $ShopeeBundle = $conn->real_escape_string($row[30]);
    $Seller_Bundle = $conn->real_escape_string($row[31]);
    $Shopee_Coins = $conn->real_escape_string($row[32]);
    $CCD_Total = $conn->real_escape_string($row[33]);
    $Product_Price_Paid_Buyer = $conn->real_escape_string($row[34]);
    $Buyer_Paid_SF = $conn->real_escape_string($row[35]);
    $Shipping_Rebate_Estimate = $conn->real_escape_string($row[36]);
    $Reverse_Shipping_Fee = $conn->real_escape_string($row[37]);
    $Service_Fee = $conn->real_escape_string($row[38]);
    $Grand_Total = $conn->real_escape_string($row[39]);
    $Estimated_SF = $conn->real_escape_string($row[40]);
    $B_Username = $conn->real_escape_string($row[41]);
    $R_Name = $conn->real_escape_string($row[42]);
    $Phone_Number = $conn->real_escape_string($row[43]);
    $Delivery_Address = $conn->real_escape_string($row[44]);
    $Town = $conn->real_escape_string($row[45]);
    $District = $conn->real_escape_string($row[46]);
    $City = $conn->real_escape_string($row[47]);
    $Province = $conn->real_escape_string($row[48]);
    $Country = $conn->real_escape_string($row[49]);
    $ZipCode = $conn->real_escape_string($row[50]);
    $Remark_Buyer = $conn->real_escape_string($row[51]);
    $Order_Complete_Time = $conn->real_escape_string($row[52]);
    $Note= $conn->real_escape_string($row[53]);
    $Invoice_Request_Type= $conn->real_escape_string($row[54]);
    $Invoice_Type= $conn->real_escape_string($row[55]);
    $Name= $conn->real_escape_string($row[56]);
    $Business_Number= $conn->real_escape_string($row[57]);
    $Address= $conn->real_escape_string($row[58]);
    $Email= $conn->real_escape_string($row[59]);

    

        // Order_ID doesn't exist, insert the new record
        $sql = "INSERT INTO tbl_shopee (
            Order_ID, 
            Order_Status, 
            Cancel_Reason, 
            Return_Refund_Status, 
            Tracking_Number, 
            Shipping_Option, 
            Shipment_Method, 
            Estimated_ShipOut_Date, 
            Ship_Time, 
            Order_Creation_Date, 
            Order_Paid_Time, 
            PSR_No, 
            Product_Name, 
            SKR_No, 
            Variation_Name, 
            Orig_Price, 
            Deal_Price, 
            Qnty, 
            Returned_Qnty, 
            Product_Subtotal, 
            Total_Discount, 
            Price_Discount, 
            Shopee_Rebate, 
            SKU_Tot_Weight, 
            Number_Items_InOrder, 
            Order_Title_Weight, 
            Seller_Voucher, 
            SAC_Cashback, 
            Shopee_Voucher, 
            Bundle_Deals_Indicator, 
            Shopee_Bundle_Discount, 
            Seller_Bundle_Discount, 
            Shopee_Coins_Offset, 
            Credit_Card_Discount_Total, 
            Product_Price_Paid_Buyer, 
            Buyer_Paid_SF, 
            Shipping_Rebate_Estimate, 
            Reverse_Shipping_Fee, 
            Service_Fee, 
            GrandTotal, 
            Estimated_SF, 
            B_Username, 
            R_Name, 
            Phone_Number, 
            Delivery_Address, 
            Town, 
            District, 
            City, 
            Province, 
            Country, 
            ZipCode, 
            Remark_Buyer, 
            Order_Complete_Time, 
            Note, 
            Invoice_Request_Type, 
            Invoice_Type, 
            FullName, 
            Business_Number, 
            MainAddress, 
            Email) VALUES (
                '$Order_ID', 
                '$Order_Status', 
                '$Cancel_Reason', 
                '$Return_Refund', 
                '$TrackingNo', 
                '$Ship_Option', 
                '$Ship_Method', 
                '$Estimated_ShipOut_Date', 
                '$ShipTime', 
                '$Order_Creation_Date', 
                '$Order_Paid_Time', 
                '$PSR_No', 
                '$Product_Name', 
                '$SKR_No', 
                '$Var_Name', 
                '$OrigPrice', 
                '$DealPrice', 
                '$Qnty', 
                '$Returned_Qnty', 
                '$Product_SubTotal', 
                '$Total_Disc', 
                '$Price_Disc_Seller', 
                '$Shopee_Rebate', 
                '$SKU_Total_Rebate', 
                '$No_Items_InOrder', 
                '$Order_Total_Weight', 
                '$Seller_Voucher', 
                '$SAC_Cashback', 
                '$Shopee_Voucher', 
                '$BundleDeals', 
                '$ShopeeBundle', 
                '$Seller_Bundle', 
                '$Shopee_Coins', 
                '$CCD_Total', 
                '$Product_Price_Paid_Buyer', 
                '$Buyer_Paid_SF', 
                '$Shipping_Rebate_Estimate', 
                '$Reverse_Shipping_Fee', 
                '$Service_Fee', 
                '$Grand_Total', 
                '$Estimated_SF', 
                '$B_Username', 
                '$R_Name', 
                '$Phone_Number', 
                '$Delivery_Address', 
                '$Town', 
                '$District', 
                '$City', 
                '$Province', 
                '$Country', 
                '$ZipCode', 
                '$Remark_Buyer', 
                '$Order_Complete_Time', 
                '$Note', 
                '$Invoice_Request_Type', 
                '$Invoice_Type', 
                '$Name', 
                '$Business_Number', 
                '$Address', 
                '$Email')";

        if ($conn->query($sql) !== TRUE) {
            echo "Error: " . $sql . "<br>" . $conn->error;
        }

    }

// Close the database connection
$conn->close();
// Send a response back to the client
$response = ['status' => 'success', 'message' => 'Data imported successfully'];
echo json_encode($response);
exit(); // Ensure that no further output is sent
?>
