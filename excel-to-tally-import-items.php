<?php
  
  // include the ExcelReader class
  include 'excel_reader.php';     

  // creates an object instance of the class, and read the excel file data
  $excel = new PhpExcelReader;

  //TODO : Specify the Excel sheet Name here
  $excel->read('item-master.xls');

  // This function returns a excel rows and columns data
  // Parameter - array with excel worksheet data
  function sheetData($sheet) {
    $re = "";
    $header = '<ENVELOPE>' .
      '<HEADER>' .
      '<TALLYREQUEST>Import Data</TALLYREQUEST>' .
      '</HEADER>' .
      '<BODY>' .
      '<IMPORTDATA>' .
      '<REQUESTDESC>' .
      '<REPORTNAME>All Masters</REPORTNAME>' .
      '</REQUESTDESC>' .
      '<REQUESTDATA>';
    $footer = '</REQUESTDATA></IMPORTDATA></BODY></ENVELOPE>';

    //It is presumed that the 1st line in the Excel sheet is the header. Hence, we read the data from row 2 onwards
    $x = 2; 
    $re .= $header;
    while($x <= $sheet['numRows']) {

      //Create UNIT Master
      $re .= "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
      $y = 1;
      while($y <= $sheet['numCols']) {
        $cell = isset($sheet['cells'][$x][$y]) ? $sheet['cells'][$x][$y] : '';
        if ($y==2) {
          //Column B of XLS
          $re .= "<UNIT NAME='$cell'><NAME>$cell</NAME><ISSIMPLEUNIT>Yes</ISSIMPLEUNIT></UNIT>";
        }
        $y++;
      } // end of inner while loop 
      $re .= "</TALLYMESSAGE>";

      //Create STOCKITEM Master
      $re .= "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
      $y = 1;
      while($y <= $sheet['numCols']) {
        $cell = isset($sheet['cells'][$x][$y]) ? $sheet['cells'][$x][$y] : '';
        if ($y==1) {
          //Column A of XLS
          $re .= "<STOCKITEM NAME='$cell'><NAME>$cell</NAME>";
        }
        if ($y==2) {
          //Column B of XLS
          $re .= "<BASEUNITS>$cell</BASEUNITS></STOCKITEM>";
        }
        $y++;
      } // end of inner while loop 
      $re .= "</TALLYMESSAGE>";


      $x++;
    } //end of while loop
    
    // ends and returns the soap request
    $re .= $footer;
    return $re;
  } // end of function


  //Main code starts here
  $excel_data = '';              // to store the data of each sheet

  // read the excel data
  $excel_data .= sheetData($excel->sheets[0]);
  
  
  //We use curl to post data to Tally  
  $ch = curl_init();
  curl_setopt($ch, CURLOPT_URL, 'http://127.0.0.1:9000/');
  curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
  curl_setopt($ch, CURLOPT_POST, true);
  curl_setopt($ch, CURLOPT_POSTFIELDS, $excel_data ); // XML data
  $response = curl_exec ($ch);
  if(curl_errno($ch)){
    echo curl_error($ch);
  }
  else{
    print "<pre lang='xml'>" . htmlspecialchars($response) . "</pre>";
    curl_close($ch);
  }  
  return "";
?>
