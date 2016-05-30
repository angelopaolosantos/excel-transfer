<?php
$errors=array(); //initialized error array
$success="";

if($_SERVER['REQUEST_METHOD'] == "POST"){ // if form is submitted, check for errors
	
	if($_FILES['file']['error']>0){
		$errors[]=$_FILES["file"]["error"];
	}
	
	if($_POST['start_row']==""){
		$errors[]="start row required";
	}
	
	if($_POST['end_row']==""){
		$errors[]="end row required";
	}
	
	if($_POST['stylecode']==""){
		$errors[]="stylecode required";
	}
	
	if($_POST['offset']==""){
		$errors[]="offset required";
	}
	
	if($_POST['reference']==""){
		$errors[]="reference required";
	}
	
	if($_POST['description']==""){
		$errors[]="description required";
	}
	
	if($_POST['pieces']==""){
		$errors[]="pieces required";
	}
	
	if($_POST['component_pointer']==""){
		$errors[]="component pointer required";
	}
	
	if($_POST['component_name']==""){
		$errors[]="component name required";
	}
	
	if($_POST['component_pieces']==""){
		$errors[]="component pieces required";
	}
	
	if($_POST['component_weight']==""){
		$errors[]="component weight required";
	}
	
	if($_POST['component_price']==""){
		$errors[]="component price required";
	}
	
	if (count($errors)==0){ // if no errors, run program
		
		//------------get values submitted
		$filename=$_FILES['file']['name'];
		$file=$_FILES['file']['tmp_name'];
		$start=$_POST['start_row'];
		$end=$_POST['end_row'];
		$stylecode_col=$_POST['stylecode'];
		$offset=intval($_POST['offset']);
		$ref_col=$_POST['reference'];
		$desc_col=$_POST['description'];
		$pieces_col=$_POST['pieces'];
		$comp_pointer=$_POST['component_pointer'];
		$comp_name_col=$_POST['component_name'];
		$comp_pieces_col=$_POST['component_pieces'];
		$comp_weight_col=$_POST['component_weight'];
		$comp_price_col=$_POST['component_price'];
		//------------end get values submitted
		
		//-----------modify php.ini settings
		ini_set("memory_limit","500M");
		ini_set('max_execution_time', 300);
		error_reporting(E_ALL);
		//-----------end modify php.ini settings
		
		//-----------RUN PROGRAM
		require_once '../Classes/PHPExcel.php'; // load excel reader script
		$objPHPExcel = new PHPExcel();
		$objReader = new PHPExcel_Reader_Excel5();
		$objReader->setReadDataOnly(true);
		$objPHPExcelWrite = new PHPExcel();
		$objPHPExcelWrite->setActiveSheetIndex(0);
		$ExcelWriteIndexRow=1; //row to start writing
		$ExcelWriteIndexColumn=22; //column to start writing components
			
		$objPHPExcel = $objReader->load($file); //load file to read
		$sheet = $objPHPExcel->getActiveSheet();
		$rowIterator = $objPHPExcel->getActiveSheet()->getRowIterator();
		
		$pointer=1;
		
		for ($x=$start;$x<$end;$x++){ // indicate starting and ending row to read
			$cell = $sheet->getCell('A'.$x);
			$check=$cell->getCalculatedValue();
			
			$y=$x; //set y as reference column
			
			if (strval($check)==strval($pointer)){ // check if cursor is on item number
				
				$cell = $sheet->getCell($ref_col.$y); 
				$reference=$cell->getCalculatedValue();
				
				$cell = $sheet->getCell($desc_col.$y);
				$description=$cell->getCalculatedValue();
				
				$cell = $sheet->getCell($pieces_col.$y);
				$qty=$cell->getCalculatedValue();
				
				$cell = $sheet->getCell($stylecode_col.($y+$offset));
				$stylecode=$cell->getCalculatedValue();
				
				$component_pointer=$y+$comp_pointer;
				$components=array();
				
				$cell = $sheet->getCell('A'.$component_pointer);
				$check=$cell->getCalculatedValue();
				
				while(strtolower($check)<>strtolower("Total")){
					$component_name="";
					$component_pcs="";
					$component_wt="";
					$component_price="";
					
					$cell = $sheet->getCell($comp_name_col.$component_pointer);
					$component_name=$cell->getCalculatedValue();
					
					$cell = $sheet->getCell($comp_pieces_col.$component_pointer);
					$component_pcs=$cell->getCalculatedValue();
					
					$cell = $sheet->getCell($comp_weight_col.$component_pointer);
					$component_wt=$cell->getCalculatedValue();
					
					$cell = $sheet->getCell($comp_price_col.$component_pointer);
					$component_price=$cell->getCalculatedValue();
					
					$component['name']=$component_name;
					$component['pcs']=$component_pcs;
					$component['wt']=$component_wt;
					$component['price']=$component_price;
					
					$components[]=$component;
					
					$component_pointer++;
					$cell=$sheet->getCell('A'.$component_pointer);
					$check=$cell->getCalculatedValue();
					
				}
				
				
				//write to excel
				$objPHPExcelWrite->getActiveSheet()->setCellValue("c".$ExcelWriteIndexRow,$stylecode);
				$objPHPExcelWrite->getActiveSheet()->setCellValue("d".$ExcelWriteIndexRow,$qty);
				$objPHPExcelWrite->getActiveSheet()->setCellValue("e".$ExcelWriteIndexRow,$description);
				$objPHPExcelWrite->getActiveSheet()->setCellValue("g".$ExcelWriteIndexRow,$reference);
				
				
				if ($components){
					foreach($components as $sil_component){
						$objPHPExcelWrite->getActiveSheet()->setCellValueByColumnAndRow($ExcelWriteIndexColumn,$ExcelWriteIndexRow, $sil_component['name']);
						$objPHPExcelWrite->getActiveSheet()->setCellValueByColumnAndRow($ExcelWriteIndexColumn+2,$ExcelWriteIndexRow, $sil_component['pcs']);
						$objPHPExcelWrite->getActiveSheet()->setCellValueByColumnAndRow($ExcelWriteIndexColumn+3,$ExcelWriteIndexRow, $sil_component['wt']);
						$objPHPExcelWrite->getActiveSheet()->setCellValueByColumnAndRow($ExcelWriteIndexColumn+4,$ExcelWriteIndexRow, $sil_component['price']);
						$ExcelWriteIndexColumn+=5;
					}
				}
				
				$ExcelWriteIndexRow++;
				$ExcelWriteIndexColumn=22;
				
				$pointer++;
				
			}
	
		}
		
		$objWriter = new PHPExcel_Writer_Excel5($objPHPExcelWrite);
		if ($objWriter->save("../Files/Billionsail/".RemoveExtension($filename)."-transfer.xls")){
			$errors[]="Failed to save file...";
		}else{
			$success='File was saved as <a href="../Files/Billionsail/'.RemoveExtension($filename).'-transfer.xls" download><em>'.RemoveExtension($filename).'-transfer.xls</em></a>';
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
	<title>Billion Sail Invoice Breakdown to Excel Transfer</title>
	<script src="js/jquery.js" ></script>
	<script src="js/bootstrap.min.js" ></script>
	<link href="css/bootstrap.min.css" rel="stylesheet">
	<link href="css/custom.css" rel="stylesheet">
	<script>
	$(function() {
		$("#default_btn").click( function()
		   {
			 $("#stylecode").attr("value","A");
			 $("#offset").attr("value","1");
			 $("#reference").attr("value","B");
			 $("#description").attr("value","C");
			 $("#pieces").attr("value","F");
			 
			 $("#component_pointer").attr("value","3");
			 $("#component_name").attr("value","A");
			 $("#component_pieces").attr("value","C");
			 $("#component_weight").attr("value","D");
			 $("#component_price").attr("value","E");
			 
			 
		   }
		);
	});
	</script>
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
		<label class="col-md-3">Starting Row</label>
		<div class="col-md-4"><input id="start_row" name="start_row" type="text" /></div>
		<p class="help-block">Row number where the 1st item is located. This is the row where the script will start reading</p>
	</div>
	<div class="form-group">
		<label class="col-md-3">Ending Row</label>
		<div class="col-md-4"><input id="end_row" name="end_row" type="text" /></div>
		<p class="help-block">Row number the script will stop reading'</p>
	</div>
	<div class="form-group">
		<label class="col-md-3">Stylecode Column</label>
		<div class="col-md-4"><input id="stylecode" name="stylecode" type="text" /></div>
		<p class="help-block">Column letter of the stylecode</p>
	</div>
	<div class="form-group">
		<label class="col-md-3">Stylecode offset</label>
		<div class="col-md-4"><input id="offset" name="offset" type="text" /></div>
		<p class="help-block">number of rows of the stylecode row from the current item pointer</p>
	</div>
	<div class="form-group">
		<label class="col-md-3">Reference Column</label>
		<div class="col-md-4"><input id="reference" name="reference" type="text" /></div>
		<p class="help-block">Column letter of the Reference</p>
	</div>
	<div class="form-group">
		<label class="col-md-3">Description Column</label>
		<div class="col-md-4"><input id="description" name="description" type="text" /></div>
		<p class="help-block">Column letter of the Description</p>
	</div>
	<div class="form-group">
		<label class="col-md-3">Pieces Column</label>
		<div class="col-md-4"><input id="pieces" type="text" name="pieces" /></div>
		<p class="help-block">Column letter of the Pieces</p>
	</div>
	<h3>Components:</h3>
	<div class="form-group">
		<label class="col-md-3">Component Offset Pointer</label>
		<div class="col-md-4"><input id="component_pointer" name="component_pointer" type="text" /></div>
		<p class="help-block">number of rows of the component row from the current item pointer</p>
	</div>
	<div class="form-group">
		<label class="col-md-3">Name Column</label>
		<div class="col-md-4"><input id="component_name" name="component_name" type="text" /></div>
		<p class="help-block">Column letter of the Component Name</p>
	</div>
	<div class="form-group">
		<label class="col-md-3">Pieces Column</label>
		<div class="col-md-4"><input id="component_pieces" name="component_pieces"  type="text" /></div>
		<p class="help-block">Column letter of the Component pieces</p>
	</div>
	<div class="form-group">
		<label class="col-md-3">Weight Column</label>
		<div class="col-md-4"><input id="component_weight" name="component_weight"  type="text" /></div>
		<p class="help-block">Column letter of the Component Weight</p>
	</div>
	<div class="form-group">
		<label class="col-md-3">Price Column</label>
		<div class="col-md-4"><input id="component_price" name="component_price" type="text" /></div>
		<p class="help-block">Column letter of the Component Price</p>
	</div>
	<div class="form-group">
		<div class="col-md-3"><input class="btn btn-default btn-lg" type="button" value="Default" id="default_btn" /></div>
		<div class="col-md-3"><input class="btn btn-default btn-lg" name="submit" type="submit" /></div>
	</div>
</form>

</body>
</html>