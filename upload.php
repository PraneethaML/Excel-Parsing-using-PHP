<?php
	
	require_once "/Library/WebServer/Documents/Classes/PHPExcel.php";

	$con=mysqli_connect("127.0.0.1","root","123456","db");
	
	//include ("Classes/PHPExcel.php");

	// Check connection
	if (mysqli_connect_errno()) {
	  echo "Failed to connect to MySQL: " . mysqli_connect_error();
	}
	
	echo "Connected to database<br>";
	
	$file = $_FILES["fileToUpload"]["tmp_name"];
	$file_name=$_FILES["fileToUpload"]["name"];
	
	$file_type = pathinfo($file_name,PATHINFO_EXTENSION);
	//echo "<br>".$file_type;
	
	$allowed = array('xlsx','xls');
	
	if(isset($_POST["submit"])) 
	{
	    if(in_array($file_type,$allowed))
		{
	    	
			echo $file_name." uploaded successfuly";
			$excelReader = PHPExcel_IOFactory::createReaderForFile($file);
			
			$html="<table border='1'>";  
			
			$excelObj = $excelReader->load($file);
			$worksheet = $excelObj->getSheet(0);
			$lastRow = $worksheet->getHighestRow();
			$lastCol = 2;
			
			for($col=0;$col<=$lastCol;$col++)
			{
				if(mysqli_real_escape_string($con,$worksheet->getCellByColumnAndRow($col, 1)->getValue())=='name')
				{
					$name_index=$col;
				}
				if(mysqli_real_escape_string($con,$worksheet->getCellByColumnAndRow($col, 1)->getValue())=='mobile')
				{
					$mobile_index=$col;	
				}
				if(mysqli_real_escape_string($con,$worksheet->getCellByColumnAndRow($col, 1)->getValue())=='email')
				{
					$email_index=$col;
				}
				
			}
			
			
			for ($row = 2; $row <= $lastRow; $row++) 
			{
				 $html.="<tr>";  
				
 				/* To display as is in browser use these variables
				
				$name_html = mysqli_real_escape_string($con,$worksheet->getCellByColumnAndRow(0, $row)->getValue());
 				$mobile_html = mysqli_real_escape_string($con,$worksheet->getCellByColumnAndRow(1, $row)->getValue());
 				$email_html = mysqli_real_escape_string($con,$worksheet->getCellByColumnAndRow(2, $row)->getValue());
				*/
				$name = mysqli_real_escape_string($con,$worksheet->getCellByColumnAndRow($name_index, $row)->getValue());
				$mobile = mysqli_real_escape_string($con,$worksheet->getCellByColumnAndRow($mobile_index, $row)->getValue());
				$email = mysqli_real_escape_string($con,$worksheet->getCellByColumnAndRow($email_index, $row)->getValue());
				
				
				$sql = "INSERT INTO excel(name, mobile, email) VALUES ('".$name."',$mobile,'".$email."')";  
				mysqli_query($con, $sql);
				$html.= '<td>'.$name.'</td>';  
				$html.= '<td>'.$mobile.'</td>';  
			    $html.= '<td>'.$email.'</td>';  
				$html.= "</tr>";  
				
				
			}
		    $html .= '</table>';  
		    echo $html;  
		
	    }
		else
		{
			echo "<br>upload only excel sheet of type .xlsx or .xls";
		}
	}
	
?>