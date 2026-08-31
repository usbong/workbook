<!--
  Copyright 2020~2026 SYSON, MICHAEL B.
  Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. You ' may obtain a copy of the License at
  http://www.apache.org/licenses/LICENSE-2.0
  Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, ' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing ' permissions and limitations under the License.
 
  @company: USBONG
  @author: SYSON, MICHAEL B.
  @date created: 20200522
  @date updated: 20260831; from 20260828
  
  Input:
  1) Summary Worksheet with counts and amounts in .csv (comma-separated value) file at the Accounting/Cashier Unit

  Output:
  1) Summary Worksheet (End Week Report) that is viewable on a Computer Web Browser  
  
  Note:
  1) We can reuse this set of instructions with other .csv files that need to be viewable on a Computer Web Browser.
  2) We can auto-generate the .csv files using Microsoft EXCEL and LibreOffice CALC.
	
  Computer Web Browser Address (Example):
  1) http://localhost/usbong_kms/server/viewSummaryReportForEndWeek.php   
-->
<?php
//defined('BASEPATH') OR exit('No direct script access allowed');
?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
 	<!-- edited by Mike, 20200811 -->
   <!-- <meta charset="utf-8"> -->
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">

    <!-- Reference: Apache Friends Dashboard index.html -->
    <!-- "Always force latest IE rendering engine or request Chrome Frame" -->
    <meta content="IE=edge,chrome=1" http-equiv="X-UA-Compatible">
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
	
    <style type="text/css">
	/**/
	                    body
                        {
							font-family: Arial;
							font-size: 12pt;


							width: 860px;
/*							
							transform: scale(0.80);
							transform-origin: 0 0;							
*/							
                        }
						
						div.copyright
						{
							text-align: center;
						}
						
						img.Image-companyLogo {
							max-width: 60%;
							height: auto;
							float: left;
							text-align: center;
							padding-left: 20px;
							padding-top: 10px;
						}

						img.Image-moscLogo {
							max-width: 20%;
							height: auto;
							float: left;
							text-align: center;
						}
						
						table.imageTable
						{
							width: 100%;
<!--							border: 1px solid #ab9c7d;		
-->
						}						

						td.tableHeaderColumn
						{
							border: 1px dotted #ab9c7d;		
							text-align: center;
						}						

						td.column
						{
							border: 1px dotted #ab9c7d;		
							text-align: left;					
						}	

						td.columnDate
						{
							width: 40px;
							height: 20px;
							
							margin: 0;
							padding: 1.5px;
							
							font-family: Arial;
							font-size: 1rem;

							text-align: right;
							
							border-bottom: 1.5px solid #444444;
							border-right: 1.5px solid #444444;
/*							
							border: 1px dotted #000000;		
*/
							border: 1px solid #000000;		

							background-color: #EDE6E6;
						}

						td.columnCount
						{
							margin: 0;
							padding: 0;
							
							font-family: Arial;
							font-size: 1rem;

							text-align: center;
							
							border-bottom: 1.5px solid #444444;
							border-right: 1.5px solid #444444;
/*							
							border: 1px dotted #000000;		
*/
							border: 1px solid #000000;		

							background-color: #EDE6E6;
						}
						
						td.columnAmtPaid
						{
							width: 40px;
							height: 20px;
							
							padding: 0;
							margin: 0;
							
							font-family: Arial;
							font-size: 1rem;

							text-align: right;

							border: 1px solid #000000;		
						}

						td.columnTotal
						{
							/*border-top: 2px solid #111111;*/
							text-align: right;
							font-weight: bold;
						}

						td.columnDirectPaymentTotal
						{
							border-top: 2px solid #111111;		
							border-bottom: 2px solid #111111; /*#ab9c7d;*/		
							text-align: right;
						}
						
						td.columnGrandTotal
						{
							border-bottom: 2px solid #111111; /*#ab9c7d;*/		
							text-align: right;
						}

						td.columnBorderBottom
						{
							border: 1px dotted #ab9c7d;		
							border-bottom: 4px double black;
							text-align: center;
						}						

						td.columnBorderBottomDotted
						{
							border: 1px dotted #ab9c7d;		
							border-bottom: 2px dotted black;
							text-align: center;
						}						

						td.columnBorderTopBottom
						{
							border: 1px dotted #ab9c7d;		
							border-top: 2px solid black;
							border-bottom: 4px double black;
							text-align: center;
						}						
						
						td.imageColumn
						{
							width: 40%;
							display: inline-block;
						}						

						td.pageNameColumn
						{
							width: 50%;
							display: inline-block;
							text-align: right;
						}			

						input.inputAnswerNum
						{
							width: 60px;
							height: 20px;
							
/*							
							margin-top: 2px;
							margin-bottom: 4px;
*/						
							margin: 0;
							padding: 0;
							
							font-family: Arial;
							font-size: 1rem;

							text-align: right;
						}

						input.inputAnswerQty
						{
							width: 30px;
							height: 20px;
							
/*							
							margin-top: 2px;
							margin-bottom: 4px;
*/						
							margin: 0;
							padding: 0;
							
							font-family: Arial;
							font-size: 1rem;

							text-align: right;
						}

						input.inputAnswerAmtPaid
						{							
							width: 75px;
							height: 20px;
						
/*							
							margin-top: 2px;
							margin-bottom: 4px;
*/						
							margin: 0;
							padding: 1.5px;

							font-family: Arial;
							font-size: 1rem;

							text-align: right;

/*														
							border-bottom: 1.5px solid #444444;
							border-right: 1.5px solid #444444;
							
							background-color: #EDE6E6;
*/							
						}

						input.inputCount
						{
							width: 20px;
							height: 20px;
							
/*							
							margin-top: 2px;
							margin-bottom: 4px;
*/						
							margin: 0;
							padding: 1.5px;
							
							font-family: Arial;
							font-size: 1rem;

							text-align: right;
							
							border-bottom: 1.5px solid #444444;
							border-right: 1.5px solid #444444;
							
							background-color: #EDE6E6;
						}

						input.inputDate
						{
							width: 100px;
							height: 20px;
							
							margin: 0;
							padding: 0;
							
							font-family: Arial;
							font-size: 1rem;

							text-align: right;
							
							border-bottom: 1.5px solid #444444;
							border-right: 1.5px solid #444444;
							
							background-color: #EDE6E6;
						}

						input.inputAnswer
						{
							width: 125px;
							height: 20px;
							
/*							
							margin-top: 2px;
							margin-bottom: 4px;
*/						
							margin: 0;
							padding: 0;
							
							font-family: Arial;
							font-size: 1rem;

							text-align: right;
						}				

						td.columnTotalLabel
						{
							font-weight: bold;
							/*background-color: #00DD00;*/
							text-align: center;
						}
						td.columnTotalLabelTwo
						{
							font-weight: bold;
							background-color: #00DD00;
							text-align: left;
						}						
    /**/
    </style>
    <title>
      Summary Report for the End Week
    </title>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <style type="text/css">
    </style>
  </head>
	  <script>
		const COUNT_COLUMN=2;
		const FEE_COLUMN=4;
		const QTY_COLUMN=5;
		const AMT_PAID_COLUMN=6;
		
		const COLUMN_COUNT_MAX=8;
		const ROW_COUNT_MAX=11;

		
		var bIsActionKeyShiftPressed=false;
			
		function onLoad() {
			bIsActionKeyShiftPressed=false;

			for (iRowCount=2; iRowCount<=8; iRowCount++) {
				processCellInput(iRowCount);
			}
			
			document.body.onkeyup = function(e){
				if (e.keyCode==16) { //key shift
				//if (e.keyCode==17) { //key control
					bIsActionKeyShiftPressed=false;
				}
			}
			
			document.body.onkeyup = function(e){
				bIsActionKeyShiftPressed=false;
			}
			
			document.body.onkeydown = function(e){
				const focusedElement = document.activeElement;

				//alert("e.keyCode: "+e.keyCode);

				//note if shift pressed with right-click; menu still appears
				if (e.keyCode==16) { //key shift
				//if (e.keyCode==17) { //key control
					bIsActionKeyShiftPressed=true;
				}

				if (bIsActionKeyShiftPressed) {
					return;
				}

				if (e.keyCode==38) { //key up
					//reference; Google AI Overview; stackoverflow
					//if active element is INPUT;
					if (focusedElement && focusedElement.tagName === "INPUT") {
						var iCurrRowIndex = Number(focusedElement.id.substring(focusedElement.id.indexOf("Id")+2,focusedElement.id.indexOf("-")));
						
						var iCurrColumnIndex = Number(focusedElement.id.substring(focusedElement.id.indexOf("-")+1));
						
						//alert(focusedElement.id);
						//alert(focusedElement.id.indexOf("Id"));
						
						//alert("iCurrRowIndex: "+iCurrRowIndex);
						//alert("iCurrColumnIndex: "+iCurrColumnIndex);
						
						iCurrRowIndex-=1;

						if (iCurrRowIndex<1) {
							iCurrRowIndex=1;
						}
						
						var cellInput = document.getElementById("cellInputId"+iCurrRowIndex+"-"+iCurrColumnIndex);

						cellInput.focus();
					}				
				}
				//else if (e.keyCode==40) { //key down
				else if ((e.keyCode==40) || (e.keyCode==13)) { //key down OR ENTER
					//reference; Google AI Overview; stackoverflow
					//if active element is INPUT;
					if (focusedElement && focusedElement.tagName === "INPUT") {
						var iCurrRowIndex = Number(focusedElement.id.substring(focusedElement.id.indexOf("Id")+2,focusedElement.id.indexOf("-")));
						
						var iCurrColumnIndex = Number(focusedElement.id.substring(focusedElement.id.indexOf("-")+1));
						
						//alert(iCurrRowIndex);
						
						iCurrRowIndex+=1;

						if (iCurrRowIndex>17) {
							iCurrRowIndex=17;
						}

						//alert(">>"iCurrRowIndex);
						
						var cellInput = document.getElementById("cellInputId"+iCurrRowIndex+"-"+iCurrColumnIndex);
						
						cellInput.focus();
					}
				}
				else if (e.keyCode==39) { //key right
					//reference; Google AI Overview; stackoverflow
					//if active element is INPUT;
					if (focusedElement && focusedElement.tagName === "INPUT") {
						var iCurrRowIndex = Number(focusedElement.id.substring(focusedElement.id.indexOf("Id")+2,focusedElement.id.indexOf("-")));
						
						var iCurrColumnIndex = Number(focusedElement.id.substring(focusedElement.id.indexOf("-")+1));
						
						//alert(focusedElement.id);
						//alert(focusedElement.id.indexOf("Id"));
						
						//alert("iCurrRowIndex: "+iCurrRowIndex);
						//alert("iCurrColumnIndex: "+iCurrColumnIndex);
						
						//iCurrRowIndex-=1;

						//CURR POSITION
						var currCellInput = document.getElementById("cellInputId"+iCurrRowIndex+"-"+iCurrColumnIndex);							
						const cursorPosition = e.target.selectionStart;
/*						
						alert("cursorPosition: "+cursorPosition);
						alert("currCellInput: "+currCellInput.value.length);
*/						
						//if (cursorPosition==currCellInput.value.length) {
						if ((cursorPosition==currCellInput.value.length) || (iCurrColumnIndex==AMT_PAID_COLUMN) || (iCurrColumnIndex==COUNT_COLUMN)) {
							//NEW POSITION
							iCurrColumnIndex+=1;

							if (iCurrColumnIndex>COLUMN_COUNT_MAX) {
								iCurrColumnIndex=COLUMN_COUNT_MAX;
							}
							
							var cellInput = document.getElementById("cellInputId"+iCurrRowIndex+"-"+iCurrColumnIndex);	
													
							cellInput.focus();
						}
					}				
				}				
				else if (e.keyCode==37) { //key left
					//reference; Google AI Overview; stackoverflow
					//if active element is INPUT;
					if (focusedElement && focusedElement.tagName === "INPUT") {
						var iCurrRowIndex = Number(focusedElement.id.substring(focusedElement.id.indexOf("Id")+2,focusedElement.id.indexOf("-")));
						
						var iCurrColumnIndex = Number(focusedElement.id.substring(focusedElement.id.indexOf("-")+1));
						
						//alert(focusedElement.id);
						//alert(focusedElement.id.indexOf("Id"));
						
						//alert("iCurrRowIndex: "+iCurrRowIndex);
						//alert("iCurrColumnIndex: "+iCurrColumnIndex);
						
						//iCurrRowIndex-=1;
						
						//CURR POSITION
						var currCellInput = document.getElementById("cellInputId"+iCurrRowIndex+"-"+iCurrColumnIndex);							
						const cursorPosition = e.target.selectionStart;
/*						
						alert("cursorPosition: "+cursorPosition);
						alert("currCellInput: "+currCellInput.value.length);
*/						
						//if (cursorPosition==0) {//currCellInput.value.length) {
						if ((cursorPosition==0) || (iCurrColumnIndex==AMT_PAID_COLUMN) || (iCurrColumnIndex==COUNT_COLUMN)) {//currCellInput.value.length) {

							//NEW POSITION
							iCurrColumnIndex-=1;

							if (iCurrColumnIndex<2) {
								iCurrColumnIndex=2;
							}
							
							var cellInput = document.getElementById("cellInputId"+iCurrRowIndex+"-"+iCurrColumnIndex);	
							
							cellInput.focus();
						}
					}				
				}	
			}			
		}

		function processCellInput(iRowCount, iColumnCount) {
			//alert(iRowCount);
			var cellInput = document.getElementById("cellInputId"+iRowCount+"-"+iColumnCount);			

			var feeCell = document.getElementById("cellInputId"+iRowCount+"-"+FEE_COLUMN);
			var qtyCell = document.getElementById("cellInputId"+iRowCount+"-"+QTY_COLUMN);
			var amtPaidCell = document.getElementById("cellInputId"+iRowCount+"-"+AMT_PAID_COLUMN);

			var grandTotal = document.getElementById("grandTotalId");

/*
			alert("iRowCount: "+iRowCount);
			alert("iColumnCount: "+iColumnCount);
*/

			if (Number.isNaN(Number(feeCell.value))) {
				feeCell.value="0";
			}
			
			if (Number.isNaN(Number(qtyCell.value))) {
				qtyCell.value="0";
			}			
			
			fOutput = (Number(feeCell.value)*Number(qtyCell.value));//.toFixed(2);
			
			amtPaidCell.value=fOutput;

/*			
			alert("fOutput: "+fOutput);
			alert(amtPaidCell.value);
*/			


			//-----------------------------
			//TOTAL PART
			//-----------------------------
			//added by Mike, 20260825
			iAmtPaidTotal=0;
			for (iRowCount=1; iRowCount<=8; iRowCount++) {
				iAmtPaidTotal += Number(document.getElementById("cellInputId"+iRowCount+"-"+AMT_PAID_COLUMN).value);
			}
			
			//alert(iAmtPaidTotal);
			
			grandTotal.innerHTML=(iAmtPaidTotal);//.toFixed(2);
			
			//return;
		}
	  </script>
  <body onload="onLoad();">
<?php
	date_default_timezone_set('Asia/Hong_Kong');
	
	$COUNT_COLUMN=2;
	$FEE_COLUMN=4;
	$QTY_COLUMN=5;
	$AMT_PAID_COLUMN=6;
	
	$COLUMN_COUNT_MAX=8;
	$ROW_COUNT_MAX=11;

	//WINDOWS machine
	if (strpos(dirname(__DIR__), ":\\")!==false) {
		//$filename="C:\\xampp\\htdocs\\usbong_newsletters\\server\\templates\\ExpensesTemplate.csv";
		$filename="templates\\ExpensesTemplate.csv";
	}
	//LINUX machine
	else {				
		$filename="./templates/ExpensesTemplate.csv";
	}
			
	date_default_timezone_set('Asia/Hong_Kong');

	$sDateToday = date("Y-m-d", strtotime(date("Y-m-d")));
	$sDateTodayTransactionFormat = date("m/d/Y", strtotime(date("Y-m-d")));
	
	echo "<table>";
	echo '<tr class="row">';

	ini_set('auto_detect_line_endings', true);

	if (!file_exists($filename)) {
		$sDateToday = Date('Y-m-d, l');

		echo $filename;

		echo "ExpensesTemplate.csv not found.";
	}
	else {
		//Reference: https://stackoverflow.com/questions/9139202/how-to-parse-a-csv-file-using-php;
		//answer by: thenetimp, 20120204T0730
		//edited by: thenetimp, 20170823T1704

		$iRowCount = -1; //we later add 1 to make start value zero (0)
		//if (($handle = fopen("test.csv", "r")) !== FALSE) {
		if (($handle = fopen($filename, "r")) !== FALSE) {
		  while (($data = fgetcsv($handle, 1000, ",")) !== FALSE) {
			$num = count($data) -1; //we add -1 for the computer to not include the excess cell due to the ending \n

		//    echo "<p> $num fields in line $row: <br /></p>\n";

			$iRowCount++;
			for ($iColumnCount=0; $iColumnCount <= $num; $iColumnCount++) {
				$cellValue = utf8_encode($data[$iColumnCount]);
				
				if (($iRowCount==0)) {
					if (($iColumnCount>=0) and ($iColumnCount<=$COLUMN_COUNT_MAX)) {
						//background color sky blue
						echo "<td class='tableHeaderColumn' bgcolor='#00A2E8'><b>".$cellValue."</b></td>";
					}
					else {
						echo "<td class='column'>".$cellValue."</td>";
					}
				}
				else if (($iRowCount>=1) && ($iRowCount<$ROW_COUNT_MAX)) {
					
						$iCurrDayOfTheWeek=date("N"); //day of the week; 7 is Sunday;

						//$sEndDate = date("Y-m-d", strtotime(date("Y-m-d")."-".$iCurrDayOfTheWeek." Day"));

						$sEndDate = date("M-d-Y", strtotime(date("M-d-Y")."-".$iCurrDayOfTheWeek." Day"));

						//echo "END: ".$sEndDate."<br/>";

						//$sStartDate = date("Y-m-d", strtotime(date("Y-m-d")."-".($iCurrDayOfTheWeek+7)." Day"));

						$sStartDate = date("M-d-Y", strtotime(date("M-d-Y")."-".($iCurrDayOfTheWeek+7)." Day"));


						//echo "START: ".$sStartDate."<br/>";

						if (($iColumnCount>=0) and ($iColumnCount<=0)) {
							echo "<td class='columnDate'>".$sStartDate."</td>";
/*							
							echo "<td class='column'>";
								echo "<input type='text' id='cellInputId".$iRowCount."-".$iColumnCount."' class='inputDate' value='".$sStartDate."' min='' max='' autofocus readonly required>";
							echo "</td>";							
*/							
						}
						else if (($iColumnCount>=1) and ($iColumnCount<=1)) {
							echo "<td class='columnDate'>".$sEndDate."</td>";
/*														
							echo "<td class='column'>";
								echo "<input type='text' id='cellInputId".$iRowCount."-".$iColumnCount."' class='inputDate' value='".$sEndDate."' min='' max='' autofocus readonly required>";
							echo "</td>";							
*/							
						}
						else if ($iColumnCount==$COUNT_COLUMN) {
/*
							echo "<td class='columnCount'>";
								echo "<input type='text' id='cellInputId".$iRowCount."-".$iColumnCount."' class='inputCount' value='".$cellValue."' min='' max='' oninput='processCellInput(".$iRowCount.",".$iColumnCount.")' autofocus readonly required>";
							echo "</td>";
*/							
							echo "<td class='columnCount'>".$cellValue."</td>";
						}								
						else if ($iColumnCount==$FEE_COLUMN) {
							echo "<td class='column'>";
								echo "<input type='text' id='cellInputId".$iRowCount."-".$iColumnCount."' class='inputAnswerNum' value='".$cellValue."' min='' max='' oninput='processCellInput(".$iRowCount.",".$iColumnCount.")' autofocus required>";
							echo "</td>";
						}
						else if ($iColumnCount==$QTY_COLUMN) {
							echo "<td class='column'>";
								echo "<input type='text' id='cellInputId".$iRowCount."-".$iColumnCount."' class='inputAnswerQty' value='".$cellValue."' min='' max='' oninput='processCellInput(".$iRowCount.",".$iColumnCount.")' autofocus required>";
							echo "</td>";
						}
						else if ($iColumnCount==$AMT_PAID_COLUMN) {
/*							
							echo "<td id='amtPaidId".$iRowCount."-".$iColumnCount."' class='columnAmtPaid'>".$cellValue."</td>";
*/							
							echo "<td class='columnAmtPaid'>";
								echo "<input type='text' id='cellInputId".$iRowCount."-".$iColumnCount."' class='inputAnswerAmtPaid' value='".$cellValue."' min='' max='' oninput='processCellInput(".$iRowCount.",".$iColumnCount.")' autofocus readonly required>";
							echo "</td>";
						}						
						else {
											
							echo "<td class='column'>";
								echo "<input type='text' id='cellInputId".$iRowCount."-".$iColumnCount."' class='inputAnswer' value='".$cellValue."' min='' max='' oninput='processCellInput(".$iRowCount.",".$iColumnCount.")' autofocus required>";
							echo "</td>";
							
							
							//echo "<td class='column'>".$cellValue."</td>";
						}
				}	
				else {
/*
					if (($iColumnCount==0) && (strpos($cellValue, "TOTAL") !== false)) {
*/
					if (($iRowCount==$ROW_COUNT_MAX) && ($iColumnCount==0)) {
						//echo "<td colspan='2' class='columnTotalLabel'>".$cellValue."</td>";
						echo "<td class='columnTotalLabel'>".$cellValue."</td>";
					}
/*
					else if (($iRowCount==$ROW_COUNT_MAX) && ($iColumnCount==1)) {
						echo "<td class='columnTotalLabelTwo'>".$cellValue."</td>";
					}
*/					
					else if (($iRowCount==$ROW_COUNT_MAX) && ($iColumnCount==$AMT_PAID_COLUMN)) {
						echo "<td id='grandTotalId' class='columnTotal'>".$cellValue."</td>";
					}
					else {
						echo "<td class='column'><b>".$cellValue."</b></td>";
					}
				}
			}
			echo '</tr><tr class="row">';
		  }
		  echo '</tr>';
		  
		  fclose($handle);
		}
	}

?>
	</table>
	<div class="copyright">
<!--
		<span>© <b>www.usbong.ph</b> 2011~<?php echo date("Y");?>. All rights reserved.</span>
-->
		<span id="disclaimerLabelId" class="spanDisclaimerLabel">
		Demo Version. Please <a class="" target="_blank" href="https://github.com/usbong/workbook">report</a> for any errors.
		</span>
	</div>		 
  </body>
</html>