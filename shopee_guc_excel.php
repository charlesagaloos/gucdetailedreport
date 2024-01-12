<?php
error_reporting(E_ALL);
ini_set('display_errors', true);

require 'vendor/autoload.php'; // Include Composer's autoloader
require 'config.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// // Create a new Spreadsheet
// $spreadsheet = new Spreadsheet();

// // Get the active sheet
// $sheet = $spreadsheet->getActiveSheet();

// Validate and sanitize input
$dateFrom = isset($_POST['dateFrom']) ? $_POST['dateFrom'] : '';
$dateTo = isset($_POST['dateTo']) ? $_POST['dateTo'] : '';
$OrderStatusFilter = isset($_POST['OrderStatusFilter']) ? $_POST['OrderStatusFilter'] : '';
$searchValue = isset($_POST['orderStatus']) ? ucwords($_POST['orderStatus']) : '';

$cancelledByBuyer = "Cancelled by buyer. Reason: Payment procedure too troublesome";
$isCancelled = "Cancelled";
$isShipping = "Shipping";
$isRefund = "Refund";

// Convert the date format (assuming the date is in the format 'F j, Y')
if (!empty($dateFrom)) {
    $dateFrom = date('Y-m-d', strtotime($dateFrom));
}

if (!empty($dateTo)) {
    $dateTo = date('Y-m-d', strtotime($dateTo));
}

// SQL query
$sql = "SELECT * FROM tbl_shopee";

// Add date range condition if $dateFrom and $dateTo are not empty
if (!empty($dateFrom) && !empty($dateTo)) {
    $sql .= " WHERE Order_Creation_Date BETWEEN '$dateFrom' AND DATE_ADD('$dateTo', INTERVAL 1 DAY)";
} elseif (!empty($dateFrom)) {
    $sql .= " WHERE Order_Creation_Date >= '$dateFrom'";
} elseif (!empty($dateTo)) {
    $sql .= " WHERE Order_Creation_Date < DATE_ADD('$dateTo', INTERVAL 1 DAY)";
} 

// Add condition for searchValue
if (!empty($searchValue)) {
    if (strpos($sql, 'WHERE') !== false) {
        // If WHERE already exists, add AND for additional condition
        $sql .= " AND Order_Status LIKE '%$searchValue%'";
    } else {
        // If WHERE doesn't exist, add WHERE for the first condition
        $sql .= " WHERE Order_Status LIKE '%$searchValue%'";
    }
}
$result = $conn->query($sql);

// Create a new Spreadsheet
$spreadsheet = new Spreadsheet();

// Get the active sheet
$sheet = $spreadsheet->getActiveSheet();

// Check if there are rows in the result
if ($result->num_rows > 0) {
    // Output data into the spreadsheet
    $rowNumber = 6; // Start populating data from row 6 (below the header)
    $currentOrderID = null;
    $currentOrderStatus = null;
    $mergeStartRow = null;
    $currentMonth = null;
    while ($row = $result->fetch_assoc()) {

        // Check if the 'Order_Paid_Time' column exists and is not empty
        if (isset($row['Order_Creation_Date']) && $row['Order_Creation_Date'] !== '') {
            
            try {
                // Create a DateTime object from the database date
                $date = new DateTime($row['Order_Creation_Date']);
                // Format the date as "F j, Y" (e.g., "September 1, 2023")
                $formattedDate = $date->format('F j, Y');


                // Check if the month has changed
                if ($date->format('F') !== $currentMonth) {
                    // Reset $rowNumber to 6 when the month changes
                    $rowNumber = 6;
                    $currentOrderID = null;
                    $currentOrderStatus = null;
                    $mergeStartRow = null;
                    $currentMonth = $date->format('F');
                }
            
                // Create a new sheet for each month
                $sheetName = $date->format('F Y');
                if (!$spreadsheet->sheetNameExists($sheetName)) {
                    $spreadsheet->createSheet()->setTitle($sheetName);
                }

                // Switch to the sheet corresponding to the current month
                $sheet = $spreadsheet->getSheetByName($sheetName);


                // Create a DateTime object from the database date
                $dateCreated = new DateTime($row['Order_Creation_Date']);
                $dateCompleted = new DateTime($row['Order_Complete_Time']);
                // Format the date as "F j, Y" (e.g., "September 1, 2023")
                $formattedDateCreated = $dateCreated->format('F j, Y');
                $formattedDateCompleted = $dateCompleted->format('F j, Y');
                $sheet->setCellValue('B' . $rowNumber, $formattedDateCreated);
                $sheet->setCellValue('C' . $rowNumber, $formattedDateCompleted);
            } catch (Exception $e) {
                // Handle the case where date parsing or formatting fails
                $sheet->setCellValue('B' . $rowNumber, 'Invalid Date');
            }
            
        } else {
            // Handle the case where 'Order_Paid_Time' is not present or empty in the result set
            $sheet->setCellValue('B' . $rowNumber, 'N/A');
        }

       // Check if the current Order ID is different from the previous one
        if ($row['Order_ID'] !== $currentOrderID || $row['Order_Status'] !== $currentOrderStatus) {
            // If different, output the Order ID and Order Status and update the current Order ID and Order Status
            $sheet->setCellValue('A' . $rowNumber, $row['Order_ID']);
            $sheet->setCellValue('B' . $rowNumber, $row['Order_Status']);
            
            // Style adjustments for Order ID
            $sheet->getStyle('A' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
            $sheet->getStyle('A' . $rowNumber)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

            // Style adjustments for Order Status
            $sheet->getStyle('B' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
            $sheet->getStyle('B' . $rowNumber)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

            // Check if there was a previous Order ID for merging
            if ($mergeStartRow !== null && $mergeStartRow !== $rowNumber - 1) {
                // Merge cells for specific columns
                $mergeColumns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC'];
                foreach ($mergeColumns as $column) {
                    $sheet->mergeCells("$column$mergeStartRow:$column" . ($rowNumber - 1));
                    $sheet->getStyle("$column$mergeStartRow:$column" . ($rowNumber - 1))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                    $sheet->getStyle("$column$mergeStartRow:$column" . ($rowNumber - 1))->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
                }
            }

            // Update variables for the current Order ID and Order Status
            $currentOrderID = $row['Order_ID'];
            $currentOrderStatus = $row['Order_Status'];
            $mergeStartRow = $rowNumber;
        }

        // Comment ko muna for no reason hahaha
        // $sheet->setCellValue('A' . $rowNumber, $row['Order_ID']);
        // $sheet->getStyle('A' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        // $sheet->getStyle('A' . $rowNumber)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        
        // Set vertical alignment for all cells in the row
        $columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC'];
        foreach ($columns as $column) {
            $sheet->getStyle($column . $rowNumber)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
            // Add border to individual cells
            $sheet->getStyle($column . $rowNumber)->getBorders()->getAllBorders()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
        }

        // Merge cells for the header
        $sheet->mergeCells('A1:AC1');
        $sheet->mergeCells('A2:AC2');
        $sheet->mergeCells('A3:AC3');
        $sheet->mergeCells('A4:J4');
        $sheet->mergeCells('K4:N4');
        $sheet->mergeCells('O4:O5');
        $sheet->mergeCells('P4:U4');
        $sheet->mergeCells('V4:V5');
        $sheet->mergeCells('X4:AC4');
    
    
        // Set background color for merged cells A1:AC1
        $sheet->getStyle('A1:AC1')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('A1:AC1')->getFill()->getStartColor()->setRGB('00FFFF'); // RGB(0, 255, 255)
    
        // Set background color for merged cells A2:AC2
        $sheet->getStyle('A2:AC2')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('A2:AC2')->getFill()->getStartColor()->setRGB('00FFFF'); // RGB(0, 255, 255)
    
        // Set background color for merged cells A3:AC3
        $sheet->getStyle('A3:AC3')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('A3:AC3')->getFill()->getStartColor()->setRGB('00FFFF'); // RGB(0, 255, 255)
    
    
    
        // Add borders to cells A1 to AC5
        $borderStyle = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN;
        $sheet->getStyle('A5:AC5')->getBorders()->getAllBorders()->setBorderStyle($borderStyle);
    
    
    
        // Set the first merged cell value and center align
        $sheet->setCellValue('A1', 'Gloire Unlimited Company');
        $sheet->getStyle('A1:AC1')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('A1:AC1')->getFont()->setBold(true)->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK))->setSize(18);
    
        // Set the second merged cell value and center align
        $sheet->setCellValue('A2', 'Shopee Net Sales');
        $sheet->getStyle('A2:AC2')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('A2:AC2')->getFont()->setBold(true)->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK))->setSize(18);
        
        foreach ($spreadsheet->getSheetNames() as $sheetName) {
            $currentSheet = $spreadsheet->getSheetByName($sheetName);
        
            // Assuming $dateFrom and $dateTo are in 'Y-m-d' format
            $startMonthYear = '';
        
            // Check if filter dates are set
            if (!empty($dateFrom) && !empty($dateTo)) {
                $startMonthYear = date('F Y', strtotime($dateFrom));
            }
        
            // Set the third merged cell value and center align
            if (!empty($startMonthYear)) {
                $currentSheet->setCellValue('A3', 'For the month of ' . $sheetName);
                $currentSheet->getStyle('A3:AC3')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $currentSheet->getStyle('A3:AC3')->getFont()->setBold(true)->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK))->setSize(16);
            }
        }
        
    
        // Set the third merged cell value and center align
        $sheet->setCellValue('A4', 'ORDER DETAILS'); 
        $sheet->getStyle('A4:J4')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);    
        // Set black background color and white font color for the range 'A4:J4'
        $sheet->getStyle('A4:J4')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('A4:J4')->getFill()->getStartColor()->setARGB('000000'); // Black color
        $sheet->getStyle('A4:J4')->getFont()->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE))->setSize(18); // White font color
        $sheet->getStyle('A4:J4')->getFont()->setBold(true); // Set text to bold
    
    
        // Set the third merged cell value and center align
        $sheet->setCellValue('K4', 'PRODUCT PRICE'); 
        $sheet->getStyle('K4:N4')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);  
        // Set dark blue background color, white font color, and make it bold for the range 'K4:N4'
        $sheet->getStyle('K4:N4')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('K4:N4')->getFill()->getStartColor()->setARGB('1F4E78'); // Dark blue color
        $sheet->getStyle('K4:N4')->getFont()->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE))->setSize(18); // White font color
        $sheet->getStyle('K4:N4')->getFont()->setBold(true); // Set text to bold
    
        // Set the background color and font color for cells A5 to J5
        $sheet->getStyle('A5:J5')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('A5:J5')->getFill()->getStartColor()->setRGB('404040'); 
        $sheet->getStyle('A5:J5')->getFont()->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE)); // White font color
    
        // Set the background color and font color for cells K5 to N5
        $sheet->getStyle('K5:N5')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('K5:N5')->getFill()->getStartColor()->setRGB('2F75B5'); // RGB 47, 117, 181
        $sheet->getStyle('K5:N5')->getFont()->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE)); // White font color
                
        // Set the third merged cell value and center align
        $sheet->setCellValue('O4', 'Product Subtotal'); 
        $sheet->getStyle('O4:O5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);    
        $sheet->getStyle('O4:O5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);   
        $sheet->getStyle('O4:O5')->getAlignment()->setWrapText(true);   
        $sheet->getStyle('O4:O5')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('O4:O5')->getFill()->getStartColor()->setARGB('000000'); // Black color
        $sheet->getStyle('O4:O5')->getFont()->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE)); // White font color
        // Set the width for column O
        $sheet->getColumnDimension('O')->setWidth(15); // Adjust the width as needed
    
        // Set the third merged cell value and center align
        $sheet->setCellValue('P4', 'DEDUCTIONS'); 
        $sheet->getStyle('P4:U4')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER); 
        // Set red background color, white font color, and make it bold for the range 'P4:U4'
        $sheet->getStyle('P4:U4')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('P4:U4')->getFill()->getStartColor()->setARGB('C00000'); // Red color
        $sheet->getStyle('P4:U4')->getFont()->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE))->setSize(18); // White font color
        $sheet->getStyle('P4:U4')->getFont()->setBold(true); // Set text to bold
    
        // Set the background color and font color for cells P5 to U5
        $sheet->getStyle('P5:U5')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('P5:U5')->getFill()->getStartColor()->setRGB('FF0000'); // Red color
        $sheet->getStyle('P5:U5')->getFont()->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE)); // White font color
    
        $sheet->setCellValue('V4', 'Net Sales'); 
        $sheet->getStyle('V4:V5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('V4:V5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getStyle('V4:V5')->getAlignment()->setWrapText(true);   
        $sheet->getStyle('V4:V5')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('V4:V5')->getFill()->getStartColor()->setARGB('000000'); // Black color
        $sheet->getStyle('V4:V5')->getFont()->setBold(true);
        $sheet->getStyle('V4:V5')->getFont()->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE)); // White font color  
        // Set the width for column V
        $sheet->getColumnDimension('V')->setWidth(15); // Adjust the width as needed
    
        // Set the third merged cell value and center align
        $sheet->setCellValue('W4', 'REFUND'); 
        $sheet->getStyle('W4')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER); 
        $sheet->getColumnDimension('W')->setWidth(15); // Adjust the width as needed
        // Set the background color, white font color, and make it bold for the cell 'W4'
        $sheet->getStyle('W4')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('W4')->getFill()->getStartColor()->setRGB('375623'); // RGB 55, 86, 35
        $sheet->getStyle('W4')->getFont()->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE))->setSize(18); // White font color
        $sheet->getStyle('W4')->getFont()->setBold(true); // Set text to bold
    
        // Set the background color and font color for cell W5
        $sheet->getStyle('W5')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('W5')->getFill()->getStartColor()->setRGB('547E35'); // RGB 84, 130, 53
        $sheet->getStyle('W5')->getFont()->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE)); // White font color
    
        // Set the third merged cell value and center align
        $sheet->setCellValue('X4', 'ADJUSTMENT'); 
        $sheet->getStyle('X4:AC4')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER); 
        // Set dark gray background color, white font color, and make it bold for the range 'X4:AC4'
        $sheet->getStyle('X4:AC4')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('X4:AC4')->getFill()->getStartColor()->setRGB('525252'); // Dark gray color
        $sheet->getStyle('X4:AC4')->getFont()->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE))->setSize(18); // White font color
        $sheet->getStyle('X4:AC4')->getFont()->setBold(true); // Set text to bold
    
    
        // Set the background color and font color for cells X5 to AC5
        $sheet->getStyle('X5:AC5')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('X5:AC5')->getFill()->getStartColor()->setRGB('7B7B7B'); // RGB 123, 123, 123
        $sheet->getStyle('X5:AC5')->getFont()->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE)); // White font color
    
        // COLUMNS
        $sheet->setCellValue('A5', 'Order ID');   
        $sheet->getStyle('A5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('A5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('A')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('A5')->getAlignment()->setWrapText(true);  
        $sheet->getStyle('A5')->getFont()->setBold(true); // Set text to bold
    
        $sheet->setCellValue('B5', 'Order Paid Date');   
        $sheet->getStyle('B5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('B5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('B')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('B5')->getAlignment()->setWrapText(true);  
        $sheet->getStyle('B5')->getFont()->setBold(true); // Set text to bold
    
        $sheet->setCellValue('C5', 'Order Complete Date');   
        $sheet->getStyle('C5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('C5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('C')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('C5')->getAlignment()->setWrapText(true);  
        $sheet->getStyle('C5')->getFont()->setBold(true); // Set text to bold
    
        $sheet->setCellValue('D5', 'Order Status');   
        $sheet->getStyle('D5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('D5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('D')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('D5')->getAlignment()->setWrapText(true);  
        $sheet->getStyle('D5')->getFont()->setBold(true); // Set text to bold
    
        $sheet->setCellValue('E5', 'Cancel Reason');   
        $sheet->getStyle('E5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('E5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('E')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('E5')->getAlignment()->setWrapText(true);  
        $sheet->getStyle('E5')->getFont()->setBold(true); // Set text to bold
    
        $sheet->setCellValue('F5', 'Failed Deliver Status');  
        $sheet->getStyle('F5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('E5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('F')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('F5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('F5')->getFont()->setBold(true); // Set text to bold
    
        $sheet->setCellValue('G5', 'Return / Refund Status');
        $sheet->getStyle('G5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('G5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('G')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('G5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('G5')->getFont()->setBold(true); // Set text to bold
    
        $sheet->setCellValue('H5', 'Tracking Number');   
        $sheet->getStyle('H5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('H5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('H')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('H5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('H5')->getFont()->setBold(true); // Set text to bold
    
        $sheet->setCellValue('I5', 'Product Name');  
        $sheet->getStyle('I5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('I5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('I')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('I5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('I5')->getFont()->setBold(true); // Set text to bold
    
        $sheet->setCellValue('J5', 'SKU Reference No.');   
        $sheet->getStyle('J5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('J5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('J')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('J5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('J5')->getFont()->setBold(true); // Set text to bold
    
        $sheet->setCellValue('K5', 'Original Price');   
        $sheet->getStyle('K5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('K5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('K')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('K5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('K5')->getFont()->setBold(true); // Set text to bold
    
        $sheet->setCellValue('L5', 'Deal Price');   
        $sheet->getStyle('L5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('L5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('L')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('L5')->getAlignment()->setWrapText(true);
        $sheet->getStyle('L5')->getFont()->setBold(true); // Set text to bold 
    
        $sheet->setCellValue('M5', 'Quantity');   
        $sheet->getStyle('M5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('M5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('M')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('M5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('M5')->getFont()->setBold(true); // Set text to bold 
    
        $sheet->setCellValue('N5', 'Total Discount');   
        $sheet->getStyle('N5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('N5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('N')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('N5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('N5')->getFont()->setBold(true); // Set text to bold 
    
        $sheet->setCellValue('P5', 'Shipping Fee Discount');   
        $sheet->getStyle('P5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('P5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('P')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('P5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('P5')->getFont()->setBold(true); // Set text to bold 
    
        $sheet->setCellValue('Q5', 'Seller Voucher');   
        $sheet->getStyle('Q5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('Q5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('Q')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('Q5')->getAlignment()->setWrapText(true);
        $sheet->getStyle('Q5')->getFont()->setBold(true); // Set text to bold  
    
        $sheet->setCellValue('R5', 'Commission Fee');   
        $sheet->getStyle('R5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('R5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('R')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('R5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('R5')->getFont()->setBold(true); // Set text to bold  
    
        $sheet->setCellValue('S5', 'Service Fee');   
        $sheet->getStyle('S5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('S5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('S')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('S5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('S5')->getFont()->setBold(true); // Set text to bold  
    
        $sheet->setCellValue('T5', 'Transaction Fee');   
        $sheet->getStyle('T5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('T5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('T')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('T5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('T5')->getFont()->setBold(true); // Set text to bold
    
        $sheet->setCellValue('U5', 'Total Deductions');   
        $sheet->getStyle('U5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('U5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('U')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('U5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('U5')->getFont()->setBold(true); // Set text to bold
    
        $sheet->setCellValue('W5', 'Refunded Amount');   
        $sheet->getStyle('W5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('W5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('W')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('W5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('W5')->getFont()->setBold(true); // Set text to bold
    
        $sheet->setCellValue('X5', 'Adjustment Complete Date');   
        $sheet->getStyle('X5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('X5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('X')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('X5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('X5')->getFont()->setBold(true); // Set text to bold
    
        $sheet->setCellValue('Y5', 'Total Adjustment Amount');   
        $sheet->getStyle('Y5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('Y5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('Y')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('Y5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('Y5')->getFont()->setBold(true); // Set text to bold
    
        $sheet->setCellValue('Z5', 'Return');   
        $sheet->getStyle('Z5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('Z5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('Z')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('Z5')->getAlignment()->setWrapText(true);
        $sheet->getStyle('Z5')->getFont()->setBold(true); // Set text to bold 
    
        $sheet->setCellValue('AA5', 'Reimbursed');   
        $sheet->getStyle('AA5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('AA5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('AA')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('AA5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('AA5')->getFont()->setBold(true); // Set text to bold 
    
        $sheet->setCellValue('AB5', 'Date Submitted');   
        $sheet->getStyle('AB5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('AB5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('AB')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('AB5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('AB5')->getFont()->setBold(true); // Set text to bold 
    
        $sheet->setCellValue('AC5', 'Status');   
        $sheet->getStyle('AC5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('AC5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getColumnDimension('AC')->setWidth(25); // Adjust the width as needed
        $sheet->getStyle('AC5')->getAlignment()->setWrapText(true); 
        $sheet->getStyle('AC5')->getFont()->setBold(true); // Set text to bold 
    
        $sheet->getRowDimension($rowNumber)->setRowHeight(50); // Adjust the height as needed

        $dateCreated = !empty($row['Order_Creation_Date']) ? new DateTime($row['Order_Creation_Date']) : null;
        $dateCompleted = !empty($row['Order_Complete_Time']) ? new DateTime($row['Order_Complete_Time']) : null;

        $formattedDateCreated = ($dateCreated !== null) ? $dateCreated->format('F j, Y') : '';
        $formattedDateCompleted = ($dateCompleted !== null) ? $dateCompleted->format('F j, Y') : '';
        
        $sheet->setCellValue('B' . $rowNumber, $formattedDateCreated);
        $sheet->getStyle('B' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);

        $sheet->setCellValue('C' . $rowNumber, $formattedDateCompleted);
        $sheet->getStyle('C' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);

        $sheet->setCellValue('D' . $rowNumber, $row['Order_Status']);
        $sheet->getStyle('D' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);

        if ($row['Order_Status'] == $isCancelled) {
            $sheet->getStyle('A' . $rowNumber . ':AC' . $rowNumber)->getFont()->getColor()->setRGB('FF0000');
        
            // Add additional condition for Cancel_Reason
            if (strpos($row['Cancel_Reason'], "Cancelled by buyer.") !== false) {
                $sheet->getStyle('A' . $rowNumber . ':AC' . $rowNumber)->getFont()->getColor()->setRGB('0070C0'); 
            }
        }elseif ($row['Order_Status'] == $isRefund) {
            $sheet->getStyle('A' . $rowNumber . ':AC' . $rowNumber)->getFont()->getColor()->setRGB('00B050'); // RGB(0, 176, 80)
        }elseif($row['Order_Status'] == $isShipping){
            $sheet->getStyle('A' . $rowNumber . ':AC' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
            $sheet->getStyle('A' . $rowNumber . ':AC' . $rowNumber)->getFill()->getStartColor()->setRGB('FFECB4'); // RGB(255, 236, 180)
        }

        $sheet->setCellValue('E' . $rowNumber, $row['Cancel_Reason']);
        $sheet->getStyle('E' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getColumnDimension('E')->setWidth(50);
        $sheet->getStyle('E' . $rowNumber)->getAlignment()->setWrapText(true); 
        
        $sheet->setCellValue('F' . $rowNumber, "");
        $sheet->getStyle('F' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);

        $sheet->setCellValue('G' . $rowNumber, $row['Return_Refund_Status']);
        $sheet->getStyle('G' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);

        $sheet->setCellValueExplicit('H' . $rowNumber, $row['Tracking_Number'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
        $sheet->getStyle('H' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);

        $sheet->setCellValue('I' . $rowNumber, $row['Product_Name']);
        $sheet->getStyle('I' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getColumnDimension('I')->setWidth(50);
        $sheet->getStyle('I' . $rowNumber)->getAlignment()->setWrapText(true); 
        $sheet->getStyle('I')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

        $sheet->setCellValue('J' . $rowNumber, $row['SKR_No']);
        $sheet->getStyle('J' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);

        $sheet->setCellValue('K' . $rowNumber, $row['Orig_Price']);
        $sheet->getStyle('K' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('K' . $rowNumber)->getNumberFormat()->setFormatCode('[$₱-421] * #,##0.00_);[Red]([$₱-421] * #,##0.00)');
        $sheet->getStyle('K' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('K' . $rowNumber)->getFill()->getStartColor()->setRGB('D9E1F2'); 

        $sheet->setCellValue('L' . $rowNumber, $row['Deal_Price']);
        $sheet->getStyle('L' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('L' . $rowNumber)->getNumberFormat()->setFormatCode('[$₱-421] * #,##0.00_);[Red]([$₱-421] * #,##0.00)');
        $sheet->getStyle('L' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('L' . $rowNumber)->getFill()->getStartColor()->setRGB('D9E1F2'); // RGB(217, 225, 242)

        $sheet->setCellValue('M' . $rowNumber, $row['Qnty']);
        $sheet->getStyle('M' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('M' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('M' . $rowNumber)->getFill()->getStartColor()->setRGB('D9E1F2'); // RGB(217, 225, 242)

        $sheet->setCellValue('N' . $rowNumber, $row['Total_Discount']);
        $sheet->getStyle('N' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('N' . $rowNumber)->getNumberFormat()->setFormatCode('[$₱-421] * #,##0.00_);[Red]([$₱-421] * #,##0.00)');
        $sheet->getStyle('N' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('N' . $rowNumber)->getFill()->getStartColor()->setRGB('D9E1F2'); // RGB(217, 225, 242)
        
        $sheet->setCellValue('O' . $rowNumber, $row['Product_Subtotal']);
        $sheet->getStyle('O' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('O' . $rowNumber)->getNumberFormat()->setFormatCode('[$₱-421] * #,##0.00_);[Red]([$₱-421] * #,##0.00)');
        $sheet->getStyle('O' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('O' . $rowNumber)->getFill()->getStartColor()->setRGB('B4C6E7'); // RGB(180, 198, 231)

        $sheet->setCellValue('P' . $rowNumber, $row['D_Shipping_Fee_Discount']);
        $sheet->getStyle('P' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('P' . $rowNumber)->getNumberFormat()->setFormatCode('[$₱-421] * #,##0.00_);[Red]([$₱-421] * #,##0.00)');
        $sheet->getStyle('P' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('P' . $rowNumber)->getFill()->getStartColor()->setRGB('FCE4D6'); // RGB(252, 228, 214)

        $sheet->setCellValue('Q' . $rowNumber, $row['Seller_Voucher']);
        $sheet->getStyle('Q' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('Q' . $rowNumber)->getNumberFormat()->setFormatCode('[$₱-421] * #,##0.00_);[Red]([$₱-421] * #,##0.00)');
        $sheet->getStyle('Q' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('Q' . $rowNumber)->getFill()->getStartColor()->setRGB('FCE4D6'); // RGB(252, 228, 214)

        $sheet->setCellValue('R' . $rowNumber, $row['D_Comission_Fee_Discount']);
        $sheet->getStyle('R' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('R' . $rowNumber)->getNumberFormat()->setFormatCode('[$₱-421] * #,##0.00_);[Red]([$₱-421] * #,##0.00)');
        $sheet->getStyle('R' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('R' . $rowNumber)->getFill()->getStartColor()->setRGB('FCE4D6'); // RGB(252, 228, 214)

        $sheet->setCellValue('S' . $rowNumber, $row['Service_Fee']);
        $sheet->getStyle('S' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('S' . $rowNumber)->getNumberFormat()->setFormatCode('[$₱-421] * #,##0.00_);[Red]([$₱-421] * #,##0.00)');
        $sheet->getStyle('S' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('S' . $rowNumber)->getFill()->getStartColor()->setRGB('FCE4D6'); // RGB(252, 228, 214)

        $sheet->setCellValue('T' . $rowNumber, $row['D_Transaction_Fee']);
        $sheet->getStyle('T' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('T' . $rowNumber)->getNumberFormat()->setFormatCode('[$₱-421] * #,##0.00_);[Red]([$₱-421] * #,##0.00)');
        $sheet->getStyle('T' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('T' . $rowNumber)->getFill()->getStartColor()->setRGB('FCE4D6'); // RGB(252, 228, 214)

        $sheet->setCellValue('U' . $rowNumber, $row['D_Total_Deductions']);
        $sheet->getStyle('U' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('U' . $rowNumber)->getNumberFormat()->setFormatCode('[$₱-421] * #,##0.00_);[Red]([$₱-421] * #,##0.00)');
        $sheet->getStyle('U' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('U' . $rowNumber)->getFill()->getStartColor()->setRGB('FCE4D6'); // RGB(252, 228, 214)

        $sheet->setCellValue('V' . $rowNumber, $row['Net_Sales']);
        $sheet->getStyle('V' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('V' . $rowNumber)->getNumberFormat()->setFormatCode('[$₱-421] * #,##0.00_);[Red]([$₱-421] * #,##0.00)');
        $sheet->getStyle('V' . $rowNumber)->getFont()->setBold(true);
        $sheet->getStyle('V' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('V' . $rowNumber)->getFill()->getStartColor()->setRGB('FFEACC'); // RGB(255, 242, 204)

        $sheet->setCellValue('W' . $rowNumber, $row['R_Refund_Amount']);
        $sheet->getStyle('W' . $rowNumber)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('W' . $rowNumber)->getNumberFormat()->setFormatCode('[$₱-421] * #,##0.00_);[Red]([$₱-421] * #,##0.00)');
        $sheet->getStyle('W' . $rowNumber)->getFont()->setBold(true);
        $sheet->getStyle('W' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('W' . $rowNumber)->getFill()->getStartColor()->setRGB('E2EFDA'); // RGB(226, 239, 218)

        $sheet->getStyle('X' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('X' . $rowNumber)->getFill()->getStartColor()->setRGB('E7E6E6'); // RGB(231, 230, 230)

        $sheet->getStyle('Y' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('Y' . $rowNumber)->getFill()->getStartColor()->setRGB('E7E6E6'); // RGB(231, 230, 230)

        $sheet->getStyle('Z' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('Z' . $rowNumber)->getFill()->getStartColor()->setRGB('E7E6E6'); // RGB(231, 230, 230)

        $sheet->getStyle('AA' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('AA' . $rowNumber)->getFill()->getStartColor()->setRGB('E7E6E6'); // RGB(231, 230, 230)

        $sheet->getStyle('AB' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('AB' . $rowNumber)->getFill()->getStartColor()->setRGB('E7E6E6'); // RGB(231, 230, 230)

        $sheet->getStyle('AC' . $rowNumber)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('AC' . $rowNumber)->getFill()->getStartColor()->setRGB('CCCCFF'); // RGB(204, 204, 255)

        $rowNumber++;

    }
    // Set vertical alignment for all cells in the sheet
    $columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC'];
    foreach ($columns as $column) {
        $sheet->getStyle("$column" . '6:' . "$column" . $rowNumber)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
    }
     // Remove the default sheet created with the workbook
     $spreadsheet->removeSheetByIndex(0);
}

// Add "Number of Total Completed" row to each sheet
foreach ($spreadsheet->getSheetNames() as $sheetName) {
    $sheet = $spreadsheet->getSheetByName($sheetName);

    // Find the last row with data in the sheet
    $lastDataRow = $sheet->getHighestDataRow() ;

    $rowComp = $lastDataRow + 3;
    $rowCan = $lastDataRow + 4;
    $rowShip = $lastDataRow + 5;
    $rowRef = $lastDataRow + 6;
    $rowTotal = $lastDataRow + 7;

    $rowPComp = $lastDataRow + 9;
    $rowPcan = $lastDataRow + 10;
    $rowPShip = $lastDataRow + 11;
    $rowPRef = $lastDataRow + 12;
    $percentage = 100;

    $rowNetSales = $lastDataRow + 2;

    // NET SALES
    $sheet->setCellValue('U' . ($lastDataRow + 2), 'Total');
    $sheet->getStyle('U' . ($lastDataRow + 2))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
   
    // Get the total net sales
    $totalNetSales = "=SUM(V6:V$lastDataRow)";
    $sheet->setCellValue('V' . ($lastDataRow + 2), $totalNetSales);
    $sheet->getStyle('V' . ($lastDataRow + 2))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('FFFF00');
    $sheet->getStyle('V' . ($lastDataRow + 2))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
    $sheet->getStyle('V' . ($lastDataRow + 2))->getNumberFormat()->setFormatCode('[$₱-421] * #,##0.00_);[Red]([$₱-421] * #,##0.00)');
    // Add a double underline to the "Total" value
    $sheet->getStyle('V' . ($lastDataRow + 2))->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOUBLE);
    $sheet->getStyle('V' . ($lastDataRow + 2))->getFont()->setBold(true);

    // TOTAL ADJUSTMENT AMOUNT
   
    $totalAdjustmentAmount = "=SUM(Y6:Y$lastDataRow)";
    
    // COMPLETED
    $sheet->setCellValue('A' . ($lastDataRow + 3), 'Number of Completed');
    $sheet->getStyle('A' . ($lastDataRow + 3))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('FCE4D6');
    $sheet->getStyle('A' . ($lastDataRow + 3))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);

    // Count the number of "Completed" in column D
    $completedCountFormula = "=COUNTIF(D6:D$lastDataRow, \"Completed\")";
    $sheet->setCellValue('B' . ($lastDataRow + 3), $completedCountFormula);
    $sheet->getStyle('B' . ($lastDataRow + 3))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('F8CBAD');
    $sheet->getStyle('B' . ($lastDataRow + 3))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);

    // Set vertical alignment for the new row
    $sheet->getStyle('A' . ($lastDataRow + 3) . ':B' . ($lastDataRow + 3))->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

    // CANCELLED
    $sheet->setCellValue('A' . ($lastDataRow + 4), 'Number of Cancelled');
    $sheet->getStyle('A' . ($lastDataRow + 4))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('FCE4D6');
    $sheet->getStyle('A' . ($lastDataRow + 4))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);
    // Count the number of "Cancelled" in column D
    $cancelledCountFormula = "=COUNTIF(D6:D$lastDataRow, \"Cancelled\")";
    $sheet->setCellValue('B' . ($lastDataRow + 4), $cancelledCountFormula);
    $sheet->getStyle('B' . ($lastDataRow + 4))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('F8CBAD');
    $sheet->getStyle('B' . ($lastDataRow + 4))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);
    // Set vertical alignment for the new rows
    $sheet->getStyle('A' . ($lastDataRow + 4) . ':B' . ($lastDataRow + 4))->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

    // SHIPPING
    $sheet->setCellValue('A' . ($lastDataRow + 5), 'Number of Shipping');
    $sheet->getStyle('A' . ($lastDataRow + 5))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('FCE4D6');
    $sheet->getStyle('A' . ($lastDataRow + 5))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);
    // Count the number of "Cancelled" in column D
    $cancelledCountFormula = "=COUNTIF(D6:D$lastDataRow, \"Shipping\")";
    $sheet->setCellValue('B' . ($lastDataRow + 5), $cancelledCountFormula);
    $sheet->getStyle('B' . ($lastDataRow + 5))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('F8CBAD');
    $sheet->getStyle('B' . ($lastDataRow + 5))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);
    // Set vertical alignment for the new rows
    $sheet->getStyle('A' . ($lastDataRow + 5) . ':B' . ($lastDataRow + 5))->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

    // REFUND
    $sheet->setCellValue('A' . ($lastDataRow + 6), 'Number of Refund');
    $sheet->getStyle('A' . ($lastDataRow + 6))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('FCE4D6');
    $sheet->getStyle('A' . ($lastDataRow + 6))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);
    // Count the number of "Cancelled" in column D
    $cancelledCountFormula = "=COUNTIF(D6:D$lastDataRow, \"Refund\")";
    $sheet->setCellValue('B' . ($lastDataRow + 6), $cancelledCountFormula);
    $sheet->getStyle('B' . ($lastDataRow + 6))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('F8CBAD');
    $sheet->getStyle('B' . ($lastDataRow + 6))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);
    // Set vertical alignment for the new rows
    $sheet->getStyle('A' . ($lastDataRow + 6) . ':B' . ($lastDataRow + 6))->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

    // TOTAL OF COMPLETED, CANCELLED, SHIPPING, REFUND
    $sheet->setCellValue('A' . ($lastDataRow + 7), 'Total');
    $sheet->getStyle('A' . ($lastDataRow + 7))->getFont()->setBold(true);
    $sheet->getStyle('A' . ($lastDataRow + 7))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('F8CBAD');
    $sheet->getStyle('A' . ($lastDataRow + 7))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);

    // Count the total number in column B
    $totalCountFormula = "=SUM(B$rowComp:B$rowRef)";
    $sheet->setCellValue('B' . ($lastDataRow + 7), $totalCountFormula);
    $sheet->getStyle('B' . ($lastDataRow + 7))->getFont()->setBold(true);
    $sheet->getStyle('B' . ($lastDataRow + 7))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('FFFF00');
    $sheet->getStyle('B' . ($lastDataRow + 7))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);

    // PERCENTAGE OF COMPLETED
    $sheet->setCellValue('A' . ($lastDataRow + 9), 'Percentage of Completed');
    $sheet->getStyle('A' . ($lastDataRow + 9))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('E2EFDA');
    $sheet->getStyle('A' . ($lastDataRow + 9))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);

    // Get the percentage of completed
    $percentageFormulaComp = "=(B$rowComp/B$rowTotal)";
    $sheet->setCellValue('B' . ($lastDataRow + 9), $percentageFormulaComp);
    $sheet->getStyle('B' . ($lastDataRow + 9))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('C6E0B4');
    $sheet->getStyle('B' . ($lastDataRow + 9))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);
    
    // Format percentage as text with two decimal places
    $sheet->getStyle('B' . ($lastDataRow + 9))->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_PERCENTAGE_00);

    // Set vertical alignment for the new rows
    $sheet->getStyle('A' . ($lastDataRow + 9) . ':B' . ($lastDataRow + 9))->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

    // PERCENTAGE OF CANCELLED
    $sheet->setCellValue('A' . ($lastDataRow + 10), 'Percentage of Cancelled');
    $sheet->getStyle('A' . ($lastDataRow + 10))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('E2EFDA');
    $sheet->getStyle('A' . ($lastDataRow + 10))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);

    // Get the percentage of cancelled
    $percentageFormulaCan = "=(B$rowCan/B$rowTotal)";
    $sheet->setCellValue('B' . ($lastDataRow + 10), $percentageFormulaCan);
    $sheet->getStyle('B' . ($lastDataRow + 10))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('C6E0B4');
    $sheet->getStyle('B' . ($lastDataRow + 10))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);
    
    // Format percentage as text with two decimal places
    $sheet->getStyle('B' . ($lastDataRow + 10))->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_PERCENTAGE_00);

    // Set vertical alignment for the new rows
    $sheet->getStyle('A' . ($lastDataRow + 10) . ':B' . ($lastDataRow + 10))->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

    // PERCENTAGE OF SHIPPING
    $sheet->setCellValue('A' . ($lastDataRow + 11), 'Percentage of Shipping');
    $sheet->getStyle('A' . ($lastDataRow + 11))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('E2EFDA');
    $sheet->getStyle('A' . ($lastDataRow + 11))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);

    // Get the percentage of Shipping
    $percentageFormulaShip = "=(B$rowShip/B$rowTotal)";
    $sheet->setCellValue('B' . ($lastDataRow + 11), $percentageFormulaShip);
    $sheet->getStyle('B' . ($lastDataRow + 11))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('C6E0B4');
    $sheet->getStyle('B' . ($lastDataRow + 11))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);
    
    // Format percentage as text with two decimal places
    $sheet->getStyle('B' . ($lastDataRow + 11))->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_PERCENTAGE_00);

    // Set vertical alignment for the new rows
    $sheet->getStyle('A' . ($lastDataRow + 11) . ':B' . ($lastDataRow + 11))->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

    // PERCENTAGE OF REFUND
    $sheet->setCellValue('A' . ($lastDataRow + 12), 'Percentage of Refund');
    $sheet->getStyle('A' . ($lastDataRow + 12))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('E2EFDA');
    $sheet->getStyle('A' . ($lastDataRow + 12))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);

    // Get the percentage of Refund
    $percentageFormulaRef = "=(B$rowRef/B$rowTotal)";
    $sheet->setCellValue('B' . ($lastDataRow + 12), $percentageFormulaRef);
    $sheet->getStyle('B' . ($lastDataRow + 12))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('C6E0B4');
    $sheet->getStyle('B' . ($lastDataRow + 12))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);
    
    // Format percentage as text with two decimal places
    $sheet->getStyle('B' . ($lastDataRow + 12))->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_PERCENTAGE_00);

    // Set vertical alignment for the new rows
    $sheet->getStyle('A' . ($lastDataRow + 12) . ':B' . ($lastDataRow + 12))->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

    // TOTAL PERCENTAGE
    $sheet->setCellValue('A' . ($lastDataRow + 13), 'Total');
    $sheet->getStyle('A' . ($lastDataRow + 13))->getFont()->setBold(true);
    $sheet->getStyle('A' . ($lastDataRow + 13))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('E2EFDA');
    $sheet->getStyle('A' . ($lastDataRow + 13))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);

    // Get the total Percentage
    $totalPercentageFormula = "=SUM(B$rowPComp:B$rowPRef)";
    $sheet->setCellValue('B' . ($lastDataRow + 13), $totalPercentageFormula);
    $sheet->getStyle('B' . ($lastDataRow + 13))->getFont()->setBold(true);
    $sheet->getStyle('B' . ($lastDataRow + 13))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('FFFF00');
    $sheet->getStyle('B' . ($lastDataRow + 13))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);

    // Format percentage as text with two decimal places
    $sheet->getStyle('B' . ($lastDataRow + 13))->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_PERCENTAGE_00);

    // Set vertical alignment for the new rows
    $sheet->getStyle('A' . ($lastDataRow + 13) . ':B' . ($lastDataRow + 13))->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

    // TOTAL NET SALES
    $sheet->setCellValue('A' . ($lastDataRow + 15), 'Total Net Sales');
    $sheet->getStyle('A' . ($lastDataRow + 15))->getFont()->setBold(true);
    $sheet->getStyle('A' . ($lastDataRow + 15))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('D9E1F2');
    $sheet->getStyle('A' . ($lastDataRow + 15))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);

    // Get the final total net sales
    $totalNetSalesFormula = "=SUM(V6:V$lastDataRow) - SUM(Y6:Y$lastDataRow)";
    $sheet->setCellValue('B' . ($lastDataRow + 15), $totalNetSalesFormula);
    $sheet->getStyle('B' . ($lastDataRow + 15))->getFont()->setBold(true);
    $sheet->getStyle('B' . ($lastDataRow + 15))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('FFFF00');
    $sheet->getStyle('B' . ($lastDataRow + 15))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
    $sheet->getStyle('B' . ($lastDataRow + 15))->getNumberFormat()->setFormatCode('[$₱-421] * #,##0.00_);[Red]([$₱-421] * #,##0.00)');
    $sheet->getStyle('B' . ($lastDataRow + 15))->getBorders()->getAllBorders()->setBorderStyle($borderStyle);

}
       
if($OrderStatusFilter == "Cancelled"){

    $excelFilename = 'Shopee_GUC_Cancelled_Report.xlsx';

}else if($OrderStatusFilter == "Completed"){

    $excelFilename = 'Shopee_GUC_Completed_Report.xlsx';

}else{
    $excelFilename = 'Shopee_GUC_Report.xlsx';
}

// Save the spreadsheet to a file
$writer = new Xlsx($spreadsheet);
$writer->save($excelFilename);

// Close the database connection
$conn->close();

// Set headers to force download
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="' . $excelFilename . '"');
header('Cache-Control: max-age=0');

// Output file to the browser
$writer->save('php://output');
exit;
?>
