<?php
$errors=array(); //initialized error array
$success="";

if($_SERVER['REQUEST_METHOD'] == "POST"){ // if form is submitted, check for errors
	
	if($_FILES['file']['error']>0){
		$errors[]=$_FILES["file"]["error"];
	}
	
	if($_POST['sheets']==""){
		$errors[]="number of sheets required";
	}
	
	if (count($errors)==0){ // if no errors, run program
		
		//------------get values submitted
		$filename=$_FILES['file']['name'];
		$file=$_FILES['file']['tmp_name'];
		$sheet_count=$_POST['sheets'];
		//------------end get values submitted
		
		//-----------modify php.ini settings
		ini_set("memory_limit","500M");
		ini_set('max_execution_time', 300);
		error_reporting(E_ALL);
		//-----------end modify php.ini settings
		
		//-----------RUN PROGRAM
		require_once '../vendor/phpoffice/phpexcel/Classes/PHPExcel.php'; // load excel reader script
		$objPHPExcel = new PHPExcel();
		$objReader = new PHPExcel_Reader_Excel5();
		$objReader->setReadDataOnly(true);
		$objPHPExcelWrite = new PHPExcel();
		$objPHPExcelWrite->setActiveSheetIndex(0);
		$ExcelWriteIndexRow=1; //row to start writing
		$ExcelWriteIndexColumn=1; //column to start writing components
			
		$objPHPExcel = $objReader->load($file); //load file to read

		$start_row=10; //row to start reading data

		for($active_sheet=0;$active_sheet<$sheet_count;$active_sheet++){ //start reading data sheets
			$objPHPExcel->setActiveSheetIndex($active_sheet);
			$sheet = $objPHPExcel->getActiveSheet();

			$col_ptr=0;
			$row_ptr=$start_row;

			$cell = $sheet->getCell('A'.$row_ptr);
			$check=$cell->getCalculatedValue();

			while(trim($check)!=""){
				for($col_ptr=0;$col_ptr<10;$col_ptr++){
					$cell_data="";
					$cell = $sheet->getCellByColumnAndRow($col_ptr,$row_ptr);
					$cell_data=$cell->getCalculatedValue();
					echo $cell_data;
					$row_data[$col_ptr]=$cell_data;
				}
				$row_datas[]=$row_data;
				$row_ptr++;
				$cell = $sheet->getCell('A'.$row_ptr);
				$check=$cell->getCalculatedValue();
			}

		}

		//Write data to new excel file
		if ($row_datas){
			$rows=count($row_datas);

			echo $rows." rows of data.";
			for($x=0;$x<$rows;$x++){
				for ($y=0;$y<10;$y++){
					$objPHPExcelWrite->getActiveSheet()->setCellValueByColumnAndRow($y,$x, $row_datas[$x][$y]);
				}
			}
		}
	
		
		$objWriter = new PHPExcel_Writer_Excel5($objPHPExcelWrite);
		if ($objWriter->save("../Files/Billionsail/".RemoveExtension($filename)."-pre.xls")){
			$errors[]="Failed to save file...";
		}else{
			$success='File was saved as <a href="../Files/Billionsail/'.RemoveExtension($filename).'-pre.xls" download><em>'.RemoveExtension($filename).'-pre.xls</em></a>';
		}
		
		//-----------END PROGRAM
	}
}

function RemoveExtension($strName) {    
	$ext = strrchr($strName, '.');    
	if($ext !== false) {       
		$strName = substr($strName, 0, -strlen($ext));    
	}
return $strName; }
?>
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!-- The above 3 meta tags *must* come first in the head; any other head content must come *after* these tags -->
    <meta name="description" content="Billion Sail Invoice Breakdown to Excel Transfer">
    <meta name="author" content="Angelo Paolo Santos">
    <link rel="icon" href="../../favicon.ico">
	<title>Billion Sail Single Sheet Excel Generator</title>
	<script src="../node_modules/jquery/dist/jquery.min.js" ></script>
	<script src="../node_modules/bootstrap/dist/js/bootstrap.min.js" ></script>
	<link href="../node_modules/bootstrap/dist/css/bootstrap.min.css" rel="stylesheet">
	<link href="../node_modules/bootstrap/dist/css/custom.css" rel="stylesheet">
</head>

<body>

<div class="container">
	<div class="row">
		<div class="col-md-12">
			<h1>Excel Transfer Generator - Billion Sail</h1>
			<h4>Select file and enter the necessary fields. </h4>
		</div>
	</div>
	<?php
	if (count($errors)>0){
		echo '<div class="alert alert-danger" role="alert">';
		echo '<h4>Something went wrong...</h4><ul>';
		
		foreach($errors as $error){
				echo "<li>".$error."</li>";
		}
		
		echo "</ul></div>";
	}
	
	if ($success){
		echo '<div class="alert alert-success" role="alert">';
		echo '<h4>Success!</h4>';
		echo $success;
		echo "</div>";
	}
	?>
</div>

<form class="form-horizontal container" action="<?php echo $_SERVER['PHP_SELF']; ?>" method="post" enctype="multipart/form-data">
	<div class="form-group">
		<label class="col-md-3">Excel File</label>
		<div class="col-md-4"><input class="form-control" id="file" name="file" type="file" /></div>
		<p class="help-block">Excel file should be save as Excel 97-2003 format</p>
	</div>
	<div class="form-group">
		<label class="col-md-3">Number of Sheets</label>
		<div class="col-md-4"><input id="sheets" name="sheets" type="text" /></div>
		<p class="help-block">Number of sheets the program will scan for data.</p>
	</div>
	<div class="form-group">
		<div class="col-md-3"><input class="btn btn-default btn-lg" name="submit" type="submit" /></div>
	</div>
</form>
</body>
</html>