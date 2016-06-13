<?php
ini_set('memory_limit', '999M');
$errors=array();
$success="";

if($_SERVER['REQUEST_METHOD'] == "POST"){
	
	if($_POST['vendor']==""){
		$errors[]="vendor required";
	}
	
	if (count($errors)==0){
		$conn = new COM("ADODB.Connection") or die("Cannot start ADO");
		$conn->Open("Provider=vfpoledb.1;Data Source=C:\wamp\www\ARZYPROGRAM\Programs\data\ARZY\jrmdbc.DBC;Collating Sequence=machine");
		
		// lets display records...
		$vendor=trim(strtolower($_POST['vendor']));
		$product_line=trim($_POST['product_line']);
		$created_from=trim($_POST['created_from']);
		$created_until=trim($_POST['created_until']);
		$contains=trim($_POST['contains']);
		$ends=trim($_POST['ends']);
		
		$onhand_qty=trim($_POST['onhand_qty']);
		
		if (isset($_POST['active_items'])){
			$active_items=trim($_POST['active_items']);
		}else{
			$active_items=0;
		}
		$argurments="";
		if (!empty($product_line)){
		 	$argurments.=" AND cprodline LIKE '%$product_line%'";
		}
		
		if (!empty($created_from)){
			$date_from=strtotime($created_from);
			$argurments.=" AND inventory.dcreate >= {".$created_from."}";
			
		}
		
		if (!empty($created_until)){
			$date_until=strtotime($created_until);
			$argurments.=" AND inventory.dcreate <= {".$created_until."}";
			
		}
		
		if (!empty($ends)){
			$date_until=strtotime($created_until);
			$argurments.=" AND inventory.cstyleno LIKE '%$ends'";
			
		}
		
		if ($active_items==1){
			$argurments.=" AND inventory.cstatus LIKE 'A'";
		}
		
		if (!empty($contains)){
			$argurments.=" AND (invitmdetail.ccode0 LIKE '%$contains%' OR inventory.ccompno0 LIKE '%$contains%' OR invitmdetail.ccode1 LIKE '%$contains%' OR inventory.ccompno1 LIKE '%$contains%' OR invitmdetail.ccode2 LIKE '%$contains%' OR inventory.ccompno2 LIKE '%$contains%' OR invitmdetail.ccode3 LIKE '%$contains%' OR inventory.ccompno3 LIKE '%$contains%' OR invitmdetail.ccode4 LIKE '%$contains%' OR inventory.ccompno4 LIKE '%$contains%' OR invitmdetail.ccode5 LIKE '%$contains%' OR inventory.ccompno5 LIKE '%$contains%' OR invitmdetail.ccode6 LIKE '%$contains%' OR inventory.ccompno6 LIKE '%$contains%' OR invitmdetail.ccode7 LIKE '%$contains%' OR inventory.ccompno7 LIKE '%$contains%' OR invitmdetail.ccode8 LIKE '%$contains%' OR inventory.ccompno8 LIKE '%$contains%' OR invitmdetail.ccode9 LIKE '%$contains%' OR inventory.ccompno9 LIKE '%$contains%' OR invitmdetail.ccode10 LIKE '%$contains%' OR inventory.ccompno10 LIKE '%$contains%' OR invitmdetail.ccode11 LIKE '%$contains%' OR inventory.ccompno11 LIKE '%$contains%' OR invitmdetail.ccode12 LIKE '%$contains%' OR inventory.ccompno12 LIKE '%$contains%' OR invitmdetail.ccode13 LIKE '%$contains%' OR inventory.ccompno13 LIKE '%$contains%' OR invitmdetail.ccode14 LIKE '%$contains%' OR inventory.ccompno14 LIKE '%$contains%' OR invitmdetail.ccode15 LIKE '%$contains%' OR inventory.ccompno15 LIKE '%$contains%')";
		}
		
		if ($onhand_qty==2){
			$argurments.=" AND invwhsqty.nonhand <= 0";
		}elseif($onhand_qty==3){
			$argurments.=" AND invwhsqty.nonhand > 0";
		}
		
		$counter_sql="SELECT COUNT(1) AS 'counter' FROM inventory LEFT JOIN invtags ON inventory.citemno=invtags.citemno LEFT JOIN invitmdetail ON inventory.citemno=invitmdetail.citemno LEFT JOIN invwhsqty on inventory.citemno=invwhsqty.citemno WHERE inventory.cvendno='$vendor'".$argurments;
		
		$rs = $conn->Execute($counter_sql); // define record set
		
		if (!$rs) {
			$errors[]=$conn->ErrorMsg();
		}else{
			$num_columns = $rs->Fields->Count();
			$fld_ctr = $rs->Fields('counter');
			
			// PERFORM READ
			
			$ma_sql="SELECT inventory.citemno,
						inventory.cbarcode,
						inventory.cstyleno,
						inventory.cdescript,
						inventory.cvendno,
						inventory.cvendmodel,
						inventory.cprodline,
						inventory.nstdcost,
						invtags.nmarkuptc,
						inventory.dcreate,
						inventory.cpathpict,
						inventory.cjobno,
						inventory.mremark,
						inventory.mnotepad,
						invitmdetail.ccode0,
						inventory.ccompno0,
						invitmdetail.ncomppcs0,
						invitmdetail.ncompqty0,	
						invitmdetail.ncompup0,
						invitmdetail.ccode1,
						inventory.ccompno1,
						invitmdetail.ncomppcs1,
						invitmdetail.ncompqty1,	
						invitmdetail.ncompup1,
						invitmdetail.ccode2,
						inventory.ccompno2,
						invitmdetail.ncomppcs2,
						invitmdetail.ncompqty2,	
						invitmdetail.ncompup2,
						invitmdetail.ccode3,
						inventory.ccompno3,
						invitmdetail.ncomppcs3,
						invitmdetail.ncompqty3,	
						invitmdetail.ncompup3,
						invitmdetail.ccode4,
						inventory.ccompno4,
						invitmdetail.ncomppcs4,
						invitmdetail.ncompqty4,	
						invitmdetail.ncompup4,
						invitmdetail.ccode5,
						inventory.ccompno5,
						invitmdetail.ncomppcs5,
						invitmdetail.ncompqty5,	
						invitmdetail.ncompup5,
						invitmdetail.ccode6,
						inventory.ccompno6,
						invitmdetail.ncomppcs6,
						invitmdetail.ncompqty6,	
						invitmdetail.ncompup6,
						invitmdetail.ccode7,
						inventory.ccompno7,
						invitmdetail.ncomppcs7,
						invitmdetail.ncompqty7,	
						invitmdetail.ncompup7,
						invitmdetail.ccode8,
						inventory.ccompno8,
						invitmdetail.ncomppcs8,
						invitmdetail.ncompqty8,	
						invitmdetail.ncompup8,
						invitmdetail.ccode9,
						inventory.ccompno9,
						invitmdetail.ncomppcs9,
						invitmdetail.ncompqty9,	
						invitmdetail.ncompup9,
						invitmdetail.ccode10,
						inventory.ccompno10,
						invitmdetail.ncomppcs10,
						invitmdetail.ncompqty10,	
						invitmdetail.ncompup10,
						invitmdetail.ccode11,
						inventory.ccompno11,
						invitmdetail.ncomppcs11,
						invitmdetail.ncompqty11,	
						invitmdetail.ncompup11,
						invitmdetail.ccode12,
						inventory.ccompno12,
						invitmdetail.ncomppcs12,
						invitmdetail.ncompqty12,	
						invitmdetail.ncompup12,
						invitmdetail.ccode13,
						inventory.ccompno13,
						invitmdetail.ncomppcs13,
						invitmdetail.ncompqty13,	
						invitmdetail.ncompup13,
						invitmdetail.ccode14,
						inventory.ccompno14,
						invitmdetail.ncomppcs14,
						invitmdetail.ncompqty14,	
						invitmdetail.ncompup14,
						invitmdetail.ccode15,
						inventory.ccompno15,
						invitmdetail.ncomppcs15,
						invitmdetail.ncompqty15,	
						invitmdetail.ncompup15 FROM inventory LEFT JOIN invtags ON inventory.citemno=invtags.citemno LEFT JOIN invitmdetail ON inventory.citemno=invitmdetail.citemno LEFT JOIN invwhsqty on inventory.citemno=invwhsqty.citemno WHERE inventory.cvendno='$vendor'".$argurments;
			
			$rs = $conn->Execute($ma_sql); // define record set
			
			if (!$rs) {
				$errors[]=$conn->ErrorMsg();
			}else{
			
				ini_set("memory_limit","500M");  // initialize excel reader
				ini_set('max_execution_time', 300);
				error_reporting(E_ALL);
			
				require_once '../vendor/autoload.php'
				
				$objPHPExcel = new PHPExcel();
				$objReader = new PHPExcel_Reader_Excel5();
				$objReader->setReadDataOnly(true);
				
				$objPHPExcelWrite = new PHPExcel(); // set excel file to write
				$objPHPExcelWrite->setActiveSheetIndex(0);
				$ExcelWriteIndexRow=1; // initialize pointers
				$ExcelWriteIndexColumn=12;
			
				$num_columns = $rs->Fields->Count();
			
			for ($i=14; $i < $num_columns; $i++) {
				$fldcomp[$i] = $rs->Fields($i);
			}
				
				$fld['itemno']=$rs->Fields('citemno');
				$fld['barcode']=$rs->Fields('cbarcode');
				$fld['stylecode']=$rs->Fields('cstyleno');
				$fld['description']=$rs->Fields('cdescript');
				$fld['vendor']=$rs->Fields('cvendno');
				$fld['model']=$rs->Fields('cvendmodel');
				$fld['product_line']=$rs->Fields('cprodline');
				$fld['standard_cost']=$rs->Fields('nstdcost');
				$fld['markup']=$rs->Fields('nmarkuptc');
				$fld['created']=$rs->Fields('dcreate');
				$fld['picture']=$rs->Fields('cpathpict');
				$fld['created_by']=$rs->Fields('cjobno');
				$fld['remark']=$rs->Fields('mremark');
				$fld['notepad']=$rs->Fields('mnotepad');
				
				while (!$rs->EOF) {
					$itemno=trim((string)$fld['itemno']->value);
					$barcode=trim((string)$fld['barcode']->value);
					$stylecode=trim((string)$fld['stylecode']->value);
					$description=trim((string)$fld['description']->value);
					$vendor=trim((string)$fld['vendor']->value);
					$model=trim((string)$fld['model']->value);
					$product_line=trim((string)$fld['product_line']->value);
					$standard_cost=trim((string)$fld['standard_cost']->value);
					$markup=trim((string)$fld['markup']->value);
					$created=trim((string)$fld['created']->value);
					$picture=trim((string)$fld['picture']->value);
					$created_by=trim((string)$fld['created_by']->value);
					$remark=trim((string)$fld['remark']->value);
					$notepad=trim((string)$fld['notepad']->value);
					
					
				
				
				$objPHPExcelWrite->getActiveSheet()->setCellValue("a".$ExcelWriteIndexRow,$itemno);
				$objPHPExcelWrite->getActiveSheet()->setCellValue("b".$ExcelWriteIndexRow,$barcode);
				$objPHPExcelWrite->getActiveSheet()->setCellValue("c".$ExcelWriteIndexRow,$stylecode);
				$objPHPExcelWrite->getActiveSheet()->setCellValue("e".$ExcelWriteIndexRow,$description);
				$objPHPExcelWrite->getActiveSheet()->setCellValue("f".$ExcelWriteIndexRow,$vendor);
				$objPHPExcelWrite->getActiveSheet()->setCellValue("g".$ExcelWriteIndexRow,$model);
				$objPHPExcelWrite->getActiveSheet()->setCellValue("i".$ExcelWriteIndexRow,$product_line);
				$objPHPExcelWrite->getActiveSheet()->setCellValue("j".$ExcelWriteIndexRow,$standard_cost);
				$objPHPExcelWrite->getActiveSheet()->setCellValue("k".$ExcelWriteIndexRow,$markup);
				$objPHPExcelWrite->getActiveSheet()->setCellValue("cq".$ExcelWriteIndexRow,$created);
				$objPHPExcelWrite->getActiveSheet()->setCellValue("cr".$ExcelWriteIndexRow,$picture);
				$objPHPExcelWrite->getActiveSheet()->setCellValue("cs".$ExcelWriteIndexRow,$created_by);
				$objPHPExcelWrite->getActiveSheet()->setCellValue("ct".$ExcelWriteIndexRow,$remark);
				$objPHPExcelWrite->getActiveSheet()->setCellValue("cu".$ExcelWriteIndexRow,$notepad);
				
				if(!empty($_POST['with_image'])){
				
					if (trim($picture)!=""){ // /PICTURES/filename.jpg
						$img_name="";
						$img_name=basename($picture);
						$image_path="./data/PICTURES/".$img_name;
					}else{
						$image_path="";
					}
					
					
					if (file_exists($image_path)){
						
						$objDrawing = new PHPExcel_Worksheet_Drawing();
						$objDrawing->setWorksheet($objPHPExcelWrite->getActiveSheet());
						$objDrawing->setName("name");
						$objDrawing->setDescription("Description");
						$objDrawing->setPath($image_path);
						$objDrawing->setCoordinates('D'.$ExcelWriteIndexRow);
						$objDrawing->setOffsetX(5);
						$objDrawing->setOffsetY(5);
						$objDrawing->setHeight(100);
					}
				}
				
				for ($i=14; $i < $num_columns; $i++) {
					$component="";
					$component=trim((string)$fldcomp[$i]->value);
					$component=($component=="0"?"":$component);
					
					$objPHPExcelWrite->getActiveSheet()->setCellValueByColumnAndRow($ExcelWriteIndexColumn,$ExcelWriteIndexRow, $component);
					$ExcelWriteIndexColumn++;
				}
				
				$ExcelWriteIndexRow++; // Move to next excel row
				$ExcelWriteIndexColumn=12; //reset column pointer
				
				$rs->MoveNext(); // Move to next Database Record Set
				
			}
			$rs->Close(); // Close connections
			$conn->Close();
			
			$objWriter = new PHPExcel_Writer_Excel5($objPHPExcelWrite);
			$fileDesc="";
			if ($_POST['vendor']!=""){
				$fileDesc.="-".trim($_POST['vendor']);
			}
			if ($_POST['product_line']!=""){
				$fileDesc.="-".trim($_POST['product_line']);
			}
			if ($_POST['created_from']!=""){
				$dateStringFrom=str_replace("/","",$_POST['created_from']);
				$fileDesc.="-".trim($dateStringFrom);
			}
			if ($_POST['created_until']!=""){
				$dateStringUntil=str_replace("/","",$_POST['created_until']);
				$fileDesc.="-".trim($dateStringUntil);
			}
			
			if($objWriter->save("../Files/Billionsail/export-".$fileDesc.".xls")){
				$errors[]="Failed to save file...";
			}else{
				$success='File was saved as <a href="../Files/Billionsail/export-'.$fileDesc.'.xls" download><em>export-'.$fileDesc.'.xls</em></a>';
			}
		}
		}
	}
}
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
	<title>Export Data from JM2000</title>
	<script src="js/jquery.js" ></script>
	<script src="js/bootstrap.min.js" ></script>
	<link href="css/bootstrap.min.css" rel="stylesheet">
	<link href="css/custom.css" rel="stylesheet">
<title>Export Data from JM2000</title>
</head>

<body>
<div class="container">
	<div class="row">
		<div class="col-md-12">
			<h1>Export Data - JM2000</h1>
			<h4>Filter data by entering the necessary fields below. </h4>
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
<label class="col-md-3">Vendor Number</label>
<div class="col-md-4"><input id="vendor" name="vendor" type="text" /></div>
</div>
<div class="form-group">
<label class="col-md-3">Product line</label>
<div class="col-md-4"><input id="product_line" name="product_line" type="text" /></div>
</div>
<div class="form-group">
<label class="col-md-3">Created from</label>
<div class="col-md-4"><input id="created_from" name="created_from" type="text" /></div>
</div>
<div class="form-group">
<label class="col-md-3">Created until</label>
<div class="col-md-4"><input id="created_until" name="created_until" type="text" /></div>
</div>
<div class="form-group">
<label class="col-md-3">Contains</label>
<div class="col-md-4"><input id="contains" name="contains" type="text" /></div>
</div>
<div class="form-group">
<label class="col-md-3">Ends with</label>
<div class="col-md-4"><input id="ends" name="ends" type="text" /></div>
</div>
<div class="form-group">
<label class="col-md-3">Active Only</label>
<div class="col-md-4"><input type="checkbox" id="active_items" name="active_items" value="1" /></div>
</div>
<div class="form-group">
<label class="col-md-3">Onhand Quantity</label>
<div class="col-md-4"><select id="onhand_qty" name="onhand_qty">
                                  <option value="1">Any</option>
                                  <option value="2">0 or less than 0</option>
                                  <option value="3">More than 0</option>
                                </select></div>
</div>
<div class="form-group">								
<label class="col-md-3">with Image</label>
<div class="col-md-4"><input type="checkbox" id="with_image" name="with_image" value="1" /></div>
</div>
<div class="form-group">
<input class="btn btn-default btn-lg" name="submit" type="submit" />
</div>
</form>
</body>
</html>