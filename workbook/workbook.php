<!--
  Copyright 2020~2026 SYSON, MICHAEL B.
  Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. You ' may obtain a copy of the License at
  http://www.apache.org/licenses/LICENSE-2.0
  Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, ' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing ' permissions and limitations under the License.
 
  @company: USBONG
  @author: SYSON, MICHAEL B.
  @date created: 20200522
  @date updated: 20260904; from 20260902
  
  Input:
  1) Expenses Template (.csv file)

  Output:
  1) Summary Worksheet (Weekly Expenses Report) that is viewable on a Computer Web Browser  
  
  Note:
  1) We can reuse this set of instructions with other .csv files that need to be viewable on a Computer Web Browser.
  2) We can auto-generate the .csv files using Microsoft EXCEL and LibreOffice CALC.
	
  Computer Web Browser Address (Example):
  1) http://localhost/usbong_kms/server/workbook/workbook.php
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

						div.divDemoVersion
						{
							margin-top: 0.5em;
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
						
						table.summaryReportTable
						{
							/*visibility: hidden;*/
							visibility: visible;
							display: none;
						}

						td.tableHeaderColumn
						{
							width: fit-content;
							height: 100%;
							
							border: 1px dotted #ab9c7d;		
							text-align: center;
						}						

						td, .column
						{
							width: fit-content;
							height: 100%;
							
							border: 1px dotted #ab9c7d;		
							border-radius: 0px;
						}	

						td.columnDate
						{
							width: 100px; /*fit-content;*/
							height: 100%;
							
							margin: 0;
							padding: 1.5px;

/*							
							font-family: monospace;
							font-size: 11pt;
*/

							text-align: center;
							
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

						td.columnWeekCount
						{
							width: fit-content;
							height: auto;
							
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
							width: 6em;
							height: auto;
							
							padding: 0;
							margin: 0;
							
							font-family: Arial;
							font-size: 1rem;

							text-align: right;

							border: 1px solid #000000;		
						}

						td.columnTotal
						{
							border: 1px dotted #ab9c7d;	
							text-align: right;
							font-weight: bold;
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
							height: 100%;
							
/*							
							margin-top: 2px;
							margin-bottom: 4px;
*/						
							margin: 0;
							padding: 0;
							
							font-family: Arial;
							font-size: 1rem;

							border-radius: 0px;
							text-align: right;
						}

						input.inputAnswerQty
						{
							width: 30px;
							height: 100%;
							
/*							
							margin-top: 2px;
							margin-bottom: 4px;
*/						
							margin: 0;
							padding: 0;
							
							font-family: Arial;
							font-size: 1rem;

							border-radius: 0px;
							text-align: right;
						}

						input.inputAnswerAmtPaid
						{							
							width: 95%;
							height: auto;
							
/*							
							margin-top: 2px;
							margin-bottom: 4px;
*/						
							margin: 0;
							padding: 0;


							font-family: Arial;
							font-size: 1rem;
							
							border: 0px;		
							border-radius: 0px;
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
							height: 100%;
							
/*							
							margin-top: 2px;
							margin-bottom: 4px;
*/						
							margin: 0;
							padding: 1.5px;
							
							font-family: Arial;
							font-size: 1rem;

							border-radius: 0px;
							text-align: right;
							
							border-bottom: 1.5px solid #444444;
							border-right: 1.5px solid #444444;
							
							background-color: #EDE6E6;
						}

						input.inputDate
						{
							width: 100px;
							height: 100%;
							
							margin: 0;
							padding: 0;
							
							font-family: Arial;
							font-size: 1rem;

							border-radius: 0px;
							text-align: right;
							
							border-bottom: 1.5px solid #444444;
							border-right: 1.5px solid #444444;
							
							background-color: #EDE6E6;
						}

						input.inputAnswer
						{
							width: 125px;
							height: 100%;
							
/*							
							margin-top: 2px;
							margin-bottom: 4px;
*/						
							margin: 0;
							padding: 0;
							
							font-family: Arial;
							font-size: 1rem;

							border-radius: 0px;
							text-align: right;
						}				

						td.columnTotalLabel
						{
							border: 1px dotted #ab9c7d;	
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

						h3
						{
							margin: 0;
							margin-bottom: 0.5em;

							padding: 0;
						}

						button.buttonSummary, .buttonWorkbook
						{
							margin: 0;
							padding-left: 0.25em;
							padding-right: 0.25em;
							padding-bottom: 0.25em;

							font-size: 12pt;
							
							border: 1px solid #ab9c7d;	
							border-radius: 2px;
							color: #000000;
						}						
    /**/
    </style>
    <title>
      Weekly Expenses Report
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
		
		const WORKBOOK_TAB=0;			
		const SUMMARY_TAB=1;
		var iCurrTab=WORKBOOK_TAB;

		var bIsActionKeyShiftPressed=false;
			
		function onLoad() {
			bIsActionKeyShiftPressed=false;
			
			var iWeekCountMax = document.getElementById("weekCountMaxId").value;				
			
			//alert(iWeekCountMax);
			
			//added by Mike, 20260904
			processWorkbook();
			
			for (iWeekCount=1; iWeekCount<iWeekCountMax; iWeekCount++) {
				for (iRowCount=1; iRowCount<ROW_COUNT_MAX; iRowCount++) {
					for (iColumnCount=1; iColumnCount<COLUMN_COUNT_MAX; iColumnCount++) {
						processCellInput(iWeekCount, iRowCount, iColumnCount);
					}
				}
			}
			
			document.body.onkeyup = function(e){
				if (e.keyCode==16) { //key shift
				//if (e.keyCode==17) { //key control
					bIsActionKeyShiftPressed=false;
				}
			}
			
			document.body.onkeydown = function(e){
				const focusedElement = document.activeElement;

				//alert("e.keyCode: "+e.keyCode);
				
				if (iCurrTab==WORKBOOK_TAB) {
				
					//cellInputId1-1-4						
					var iWeekCount = Number(focusedElement.id.substring(focusedElement.id.indexOf("Id")+2,focusedElement.id.indexOf("-")));

					var iCurrRowIndex = Number(focusedElement.id.substring(focusedElement.id.indexOf("-")+1,focusedElement.id.indexOf("-",focusedElement.id.indexOf("-")+1)));

					var iCurrColumnIndex = Number(focusedElement.id.substring(focusedElement.id.indexOf("-",focusedElement.id.indexOf("-")+1)+1));			
									
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
	/*
							alert("iWeekCount: "+iWeekCount);
							alert("iCurrRowIndex: "+iCurrRowIndex);
							alert("iCurrColumnIndex: "+iCurrColumnIndex);
	*/
							
							iCurrRowIndex-=1;

							if (iCurrRowIndex<1) {
								iCurrRowIndex=1;
							}
							
							var cellInput = document.getElementById("cellInputId"+iWeekCount+"-"+iCurrRowIndex+"-"+iCurrColumnIndex);

							cellInput.focus();
						}				
					}
					//else if (e.keyCode==40) { //key down
					else if ((e.keyCode==40) || (e.keyCode==13)) { //key down OR ENTER
						//reference; Google AI Overview; stackoverflow
						//if active element is INPUT;
						if (focusedElement && focusedElement.tagName === "INPUT") {
							iCurrRowIndex+=1;
						
							if (iCurrRowIndex>ROW_COUNT_MAX) {
								iCurrRowIndex=ROW_COUNT_MAX;
							}

							//alert(">>"iCurrRowIndex);
							
							var cellInput = document.getElementById("cellInputId"+iWeekCount+"-"+iCurrRowIndex+"-"+iCurrColumnIndex);
							
							cellInput.focus();
						}
					}
					else if (e.keyCode==39) { //key right
						//reference; Google AI Overview; stackoverflow
						//if active element is INPUT;
						if (focusedElement && focusedElement.tagName === "INPUT") {
							//CURR POSITION
							var currCellInput = document.getElementById("cellInputId"+iWeekCount+"-"+iCurrRowIndex+"-"+iCurrColumnIndex);							
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
								
								var cellInput = document.getElementById("cellInputId"+iWeekCount+"-"+iCurrRowIndex+"-"+iCurrColumnIndex);	
														
								cellInput.focus();
							}
						}				
					}				
					else if (e.keyCode==37) { //key left
						//reference; Google AI Overview; stackoverflow
						//if active element is INPUT;
						if (focusedElement && focusedElement.tagName === "INPUT") {
							//CURR POSITION
							var currCellInput = document.getElementById("cellInputId"+iWeekCount+"-"+iCurrRowIndex+"-"+iCurrColumnIndex);							
							const cursorPosition = e.target.selectionStart;

							//if (cursorPosition==0) {//currCellInput.value.length) {
							if ((cursorPosition==0) || (iCurrColumnIndex==AMT_PAID_COLUMN) || (iCurrColumnIndex==COUNT_COLUMN)) {//currCellInput.value.length) {

								//NEW POSITION
								iCurrColumnIndex-=1;

								if (iCurrColumnIndex<2) {
									iCurrColumnIndex=2;
								}
								
								var cellInput = document.getElementById("cellInputId"+iWeekCount+"-"+iCurrRowIndex+"-"+iCurrColumnIndex);	
								
								cellInput.focus();
							}
						}				
					}	
				}			
				else {
					//alert("HERE!!");
					//var summaryGrandTotalAvePerWeekColumn = document.getElementById("summaryGrandTotalAvePerWeekColumnId");

					iSummaryWeekCount=Number(focusedElement.id.substring(focusedElement.id.indexOf("Id")+2));
					
					if (e.keyCode==38) { //key up
						//reference; Google AI Overview; stackoverflow
						//if active element is INPUT;
						if (focusedElement && focusedElement.tagName === "INPUT") {
							if (focusedElement.id=="summaryGrandTotalAvePerWeekColumnId") {
								return;
							}
							
	/*
							alert("iWeekCount: "+iWeekCount);
							alert("iCurrRowIndex: "+iCurrRowIndex);
							alert("iCurrColumnIndex: "+iCurrColumnIndex);
	*/
							
							//alert("focusedElement.id"+focusedElement.id);
							//alert(iSummaryWeekCount);
							
							iSummaryWeekCount-=1;

							if (iSummaryWeekCount<1) {
								iSummaryWeekCount=1;
							}
							
							var cellInput = document.getElementById("summaryTotalColumnId"+iSummaryWeekCount);

							cellInput.focus();
						}				
					}
					//else if (e.keyCode==40) { //key down
					else if ((e.keyCode==40) || (e.keyCode==13)) { //key down OR ENTER
						//reference; Google AI Overview; stackoverflow
						//if active element is INPUT;
						if (focusedElement && focusedElement.tagName === "INPUT") {
							if (focusedElement.id=="summaryGrandTotalAvePerWeekColumnId") {
								return;
							}
							
							var cellInput = document.getElementById("cellInputId"+iWeekCount+"-"+iCurrRowIndex+"-"+iCurrColumnIndex);
							
							iSummaryWeekCount+=1;

							if (iSummaryWeekCount>iWeekCountMax) {
								iSummaryWeekCount=iWeekCountMax;
							}
							
							var cellInput = document.getElementById("summaryTotalColumnId"+iSummaryWeekCount);							
							cellInput.focus();
						}
					}
				}
			}

/*			
			//buggy
			document.body.addEventListener('touchend', (event) => {	
			//document.body.addEventListener('pointerup', (event) => {	
				const focusedElement = document.activeElement;

				//alert("e.keyCode: "+e.keyCode);
				switch(event.pointerType) {
					//case 'mouse':
					case 'touch': //displayed twice on iPad; also not when the focusedElement is touched;
						if (iCurrTab==WORKBOOK_TAB) {
						}
						else {
							if (focusedElement.id=="summaryGrandTotalAvePerWeekColumnId") {
								alert("THIS IS THE WEEKLY AVERAGE.");
								//return;
							}
						}
						break;
				}
			});
*/			
		}

		function processCellInput(iWeekCount, iRowCount, iColumnCount) {
			//alert(iRowCount);
			var cellInput = document.getElementById("cellInputId"+iWeekCount+"-"+iRowCount+"-"+iColumnCount);			
			var feeCell = document.getElementById("cellInputId"+iWeekCount+"-"+iRowCount+"-"+FEE_COLUMN);
			var qtyCell = document.getElementById("cellInputId"+iWeekCount+"-"+iRowCount+"-"+QTY_COLUMN);
			var amtPaidCell = document.getElementById("cellInputId"+iWeekCount+"-"+iRowCount+"-"+AMT_PAID_COLUMN);

			var grandTotal = document.getElementById("grandTotalId"+iWeekCount);

/*
			alert("iRowCount: "+iRowCount);
			alert("iColumnCount: "+iColumnCount);
*/

			if (Number.isNaN(Number(feeCell.value))) {
				feeCell.value="0.00";
			}
			
			if (Number.isNaN(Number(qtyCell.value))) {
				qtyCell.value="0";
			}			
			
			fOutput = (Number(feeCell.value)*Number(qtyCell.value));//.toFixed(2);
			
			amtPaidCell.value=fOutput.toFixed(2);

			//removed by Mike, 20260904
			//feeCell.value=Number(feeCell.value).toFixed(2);

/*			
			alert("fOutput: "+fOutput);
			alert(amtPaidCell.value);
*/			


			//-----------------------------
			//TOTAL PART
			//-----------------------------
			//added by Mike, 20260825
			fAmtPaidTotal=0;
			for (iRowCount=1; iRowCount<ROW_COUNT_MAX; iRowCount++) {
				fAmtPaidTotal += Number(document.getElementById("cellInputId"+iWeekCount+"-"+iRowCount+"-"+AMT_PAID_COLUMN).value);
			}
			
			//alert(iAmtPaidTotal);
			
			grandTotal.innerHTML=(fAmtPaidTotal).toFixed(2);
			
			//return;
		}		
		
		function processSummary() {
			var tableSummaryReport = document.getElementById("tableSummaryReportId");
			var iWeekCountMax = document.getElementById("weekCountMaxId").value;	
			
			var fGrandTotal=0;
			var fGrandTotalAvePerWeek=0;
			
			iCurrTab=SUMMARY_TAB;

			for (iWeekCount=1; iWeekCount<iWeekCountMax; iWeekCount++) {
				var tableWeeklyExpensesReport = document.getElementById("tableWeeklyExpensesReportId"+iWeekCount);
				var grandTotal = document.getElementById("grandTotalId"+iWeekCount);
				var summaryTotalColumn = document.getElementById("summaryTotalColumnId"+iWeekCount);
				var summaryGrandTotalColumn = document.getElementById("summaryGrandTotalColumnId");
				var summaryGrandTotalAvePerWeekColumn = document.getElementById("summaryGrandTotalAvePerWeekColumnId");
				
				//alert(grandTotal.innerHTML);
				//alert(summaryTotalColumn.value);

				summaryTotalColumn.value=grandTotal.innerHTML;
				
				fGrandTotal+=Number(grandTotal.innerHTML);
				
				//tableWeeklyExpensesReport.style.visibility="hidden";
				tableWeeklyExpensesReport.style.display="none";
			}
			
			fGrandTotalAvePerWeek=fGrandTotal/(iWeekCountMax-1);
			
			//alert(fGrandTotal.toFixed(2));
			summaryGrandTotalColumn.innerHTML=fGrandTotal.toFixed(2);
			//summaryGrandTotalAvePerWeekColumn.innerHTML=fGrandTotalAvePerWeek.toFixed(2);
			summaryGrandTotalAvePerWeekColumn.value=fGrandTotalAvePerWeek.toFixed(2);

			//tableSummaryReport.style.visibility="visible";
			tableSummaryReport.style.display="inline-block";
		}
		
		function processWorkbook() {
			var tableSummaryReport = document.getElementById("tableSummaryReportId");
			var iWeekCountMax = document.getElementById("weekCountMaxId").value;	

			var fColumnDateWidthMax=0;
			
			iCurrTab=WORKBOOK_TAB;

			for (iWeekCount=1; iWeekCount<iWeekCountMax; iWeekCount++) {
				var tableWeeklyExpensesReport = document.getElementById("tableWeeklyExpensesReportId"+iWeekCount);

				var columnStartDate = document.getElementById("columnStartDateId"+iWeekCount);
				var columnEndDate = document.getElementById("columnEndDateId"+iWeekCount);
				
				//alert(fColumnStartDate);
				//alert(document.getElementById("columnStartDateId"+iWeekCount));
				
				//tableWeeklyExpensesReport.style.visibility="visible";
				tableWeeklyExpensesReport.style.display="inline-block";
				
				//tableSummaryReport.style.visibility="hidden";
				tableSummaryReport.style.display="none";

				//-----------------
				// auto-set column width; for date column; part 1
				//-----------------

				//do this after setting tableWeeklyExpensesReport's display to inline-block
				//alert(columnStartDate.offsetWidth);
				
				var fColumnDateWidthMaxTemp=parseFloat(columnStartDate.offsetWidth);

				//alert(fColumnDateWidthMaxTemp);
				//alert(parseFloat(columnEndDate.offsetWidth));

				if (fColumnDateWidthMaxTemp<parseFloat(columnEndDate.offsetWidth)) {
					fColumnDateWidthMaxTemp=parseFloat(columnEndDate.offsetWidth);
				}

				//alert("fColumnDateWidthMaxTemp: "+fColumnDateWidthMaxTemp);

				if (fColumnDateWidthMax<fColumnDateWidthMaxTemp) {
					fColumnDateWidthMax=fColumnDateWidthMaxTemp;
				}
			}

			//-----------------
			// auto-set column width; for date column; part 2
			//-----------------

			//alert(fColumnDateWidthMax);

			for (iWeekCount=1; iWeekCount<iWeekCountMax; iWeekCount++) {
				var columnStartDate = document.getElementById("columnStartDateId"+iWeekCount);
				var columnEndDate = document.getElementById("columnEndDateId"+iWeekCount);
				
				columnStartDate.style.width=fColumnDateWidthMax+"px"; //100+"px";
				columnEndDate.style.width=fColumnDateWidthMax+"px"; //100+"px";
			}
		}
	  </script>
  <body onload="onLoad();">
  <h3>Weekly Expenses Report | <button class="buttonSummary" onclick="processSummary();">Summary</button> <button class="buttonWorkbook" onclick="processWorkbook();">Workbook</button></h3>
    
<?php
	//------------------------------------
	// Workbook
	//------------------------------------
	
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

	$iWeekCount=0;
	
	//----------------------------
	$iCurrDayOfTheWeek=date("N"); //day of the week; 7 is Sunday;

	$sCurrDateYear = date("Y");
	//echo $sCurrDateYear."<br/>";
	//echo (intval($sCurrDateYear)-1)."<br/>";

	$iCurrYear=intval($sCurrDateYear);
	$iPrevYear=intval($sCurrDateYear)-1;

	do {
		$iWeekCount++;
		
		//edited by Mike, 20260902
		//$sEndDate = date("M-d-Y", strtotime(date("M-d-Y")."-".$iCurrDayOfTheWeek." Day"));

		$sEndDate = date("M-d-Y", strtotime(date("M-d-Y")."-".($iCurrDayOfTheWeek+(7*($iWeekCount-1)))." Day"));
		
		$sEndDateYear = date("Y", strtotime(date("M-d-Y")."-".($iCurrDayOfTheWeek+(7*($iWeekCount-1)))." Day"));
		
		$iEndDateYear=intval($sEndDateYear);
		
		//echo $iEndDateYear;
		
		//$sEndYear = date("Y", strtotime(date("M-d-Y")."-".($iCurrDayOfTheWeek+(7*($iWeekCount-1)))." Day"));

		//echo "END: ".$sEndDate."<br/>";
		//echo "END: ".$sEndYear."<br/>";

		//$sStartDate = date("Y-m-d", strtotime(date("Y-m-d")."-".($iCurrDayOfTheWeek+7)." Day"));

		//$sStartDate = date("M-d-Y", strtotime(date("M-d-Y")."-".($iCurrDayOfTheWeek+6)." Day"));

		$sStartDate = date("M-d-Y", strtotime(date("M-d-Y")."-".($iCurrDayOfTheWeek+(7*$iWeekCount-1))." Day"));
		
		$sStartDateYear = date("Y", strtotime(date("M-d-Y")."-".($iCurrDayOfTheWeek+(7*$iWeekCount-1))." Day"));
		
		$iStartDateYear=intval($sStartDateYear);
		//----------------------------
		
		echo "<table id='tableWeeklyExpensesReportId".$iWeekCount."'>";
		echo '<tr class="row">';

		ini_set('auto_detect_line_endings', true);

		if (!file_exists($filename)) {
			//$sDateToday = Date('Y-m-d, l');

			echo $filename." not found.";

			//echo "ExpensesTemplate.csv not found.";
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
						
	/*					
							$iCurrDayOfTheWeek=date("N"); //day of the week; 7 is Sunday;

							//$sEndDate = date("Y-m-d", strtotime(date("Y-m-d")."-".$iCurrDayOfTheWeek." Day"));

							$sEndDate = date("M-d-Y", strtotime(date("M-d-Y")."-".$iCurrDayOfTheWeek." Day"));

							//echo "END: ".$sEndDate."<br/>";

							//$sStartDate = date("Y-m-d", strtotime(date("Y-m-d")."-".($iCurrDayOfTheWeek+7)." Day"));

							$sStartDate = date("M-d-Y", strtotime(date("M-d-Y")."-".($iCurrDayOfTheWeek+7)." Day"));
	*/

							//echo "START: ".$sStartDate."<br/>";

							if (($iColumnCount>=0) and ($iColumnCount<=0)) {
								echo "<td id='columnStartDateId".$iWeekCount."' class='columnDate'>".$sStartDate."</td>";
	/*							
								echo "<td class='column'>";
									echo "<input type='text' id='cellInputId".$iWeekCount."-".$iRowCount."-".$iColumnCount."' class='inputDate' value='".$sStartDate."' min='' max=''  readonly required>";
								echo "</td>";							
	*/							
							}
							else if (($iColumnCount>=1) and ($iColumnCount<=1)) {
								echo "<td id='columnEndDateId".$iWeekCount."' class='columnDate'>".$sEndDate."</td>";
	/*														
								echo "<td class='column'>";
									echo "<input type='text' id='cellInputId".$iWeekCount."-".$iRowCount."-".$iColumnCount."' class='inputDate' value='".$sEndDate."' min='' max=''  readonly required>";
								echo "</td>";							
	*/							
							}
							else if ($iColumnCount==$COUNT_COLUMN) {
	/*
								echo "<td class='columnCount'>";
									echo "<input type='text' id='cellInputId".$iWeekCount."-".$iRowCount."-".$iColumnCount."' class='inputCount' value='".$cellValue."' min='' max='' oninput='processCellInput(".$iWeekCount.",".$iRowCount.",".$iColumnCount.")'  readonly required>";
								echo "</td>";
	*/							
								echo "<td class='columnCount'>".$cellValue."</td>";
							}								
							else if ($iColumnCount==$FEE_COLUMN) {
								echo "<td class='column'>";
									echo "<input type='text' id='cellInputId".$iWeekCount."-".$iRowCount."-".$iColumnCount."' class='inputAnswerNum' value='".$cellValue."' min='' max='' oninput='processCellInput(".$iWeekCount.",".$iRowCount.",".$iColumnCount.")'  required>";
								echo "</td>";
							}
							else if ($iColumnCount==$QTY_COLUMN) {
								echo "<td class='column'>";
									echo "<input type='text' id='cellInputId".$iWeekCount."-".$iRowCount."-".$iColumnCount."' class='inputAnswerQty' value='".$cellValue."' min='' max='' oninput='processCellInput(".$iWeekCount.",".$iRowCount.",".$iColumnCount.")'  required>";
								echo "</td>";
							}
							else if ($iColumnCount==$AMT_PAID_COLUMN) {
	/*							
								echo "<td id='amtPaidId".$iRowCount."-".$iColumnCount."' class='columnAmtPaid'>".$cellValue."</td>";
	*/							

								echo "<td class='columnAmtPaid'>";
									echo "<input type='text' id='cellInputId".$iWeekCount."-".$iRowCount."-".$iColumnCount."' class='inputAnswerAmtPaid' value='".$cellValue."' min='' max='' oninput='processCellInput(".$iWeekCount.",".$iRowCount.",".$iColumnCount.")'  readonly required>";
								echo "</td>";
	/*
							echo "<td id='cellInputId".$iWeekCount."-".$iRowCount."-".$iColumnCount."' class='inputAnswerAmtPaid'>".$cellValue."</td>";	
	*/						
							}						
							else {
												
								echo "<td class='column'>";
									echo "<input type='text' id='cellInputId".$iWeekCount."-".$iRowCount."-".$iColumnCount."' class='inputAnswer' value='".$cellValue."' min='' max='' oninput='processCellInput(".$iWeekCount.",".$iRowCount.",".$iColumnCount.")'  required>";
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
							echo "<td id='grandTotalId".$iWeekCount."' class='columnTotal'>".$cellValue."</td>";
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
	}
	while (($iEndDateYear!=$iPrevYear) && ($iStartDateYear!=$iPrevYear));
	
	//echo $iWeekCount;
	
	$iWeekCount++;
?>

<?php
	//------------------------------------
	// Summary
	//------------------------------------
	
	date_default_timezone_set('Asia/Hong_Kong');
	
	$SUMMARY_WEEK_COUNT_COLUMN=0;
	$SUMMARY_START_DATE_COLUMN=1;
	$SUMMARY_END_DATE_COLUMN=2;
	$SUMMARY_TOTAL_COLUMN=3;
	
	$SUMMARY_COLUMN_COUNT_MAX=4;
	$SUMMARY_ROW_COUNT_MAX=2;//$iWeekCount;
	
	//WINDOWS machine
	if (strpos(dirname(__DIR__), ":\\")!==false) {
		//$filename="C:\\xampp\\htdocs\\usbong_newsletters\\server\\templates\\ExpensesTemplate.csv";
		$filename="templates\\ExpensesTemplateTotal.csv";
	}
	//LINUX machine
	else {				
		$filename="./templates/ExpensesTemplateTotal.csv";
	}
			
	date_default_timezone_set('Asia/Hong_Kong');

	$sDateToday = date("Y-m-d", strtotime(date("Y-m-d")));
	$sDateTodayTransactionFormat = date("m/d/Y", strtotime(date("Y-m-d")));

	$iSummaryWeekCount=0;
	
	//----------------------------
	$iCurrDayOfTheWeek=date("N"); //day of the week; 7 is Sunday;

	$sCurrDateYear = date("Y");
	//echo $sCurrDateYear."<br/>";
	//echo (intval($sCurrDateYear)-1)."<br/>";

	$iCurrYear=intval($sCurrDateYear);
	$iPrevYear=intval($sCurrDateYear)-1;

	do {
		$iSummaryWeekCount++;
		
		//edited by Mike, 20260902
		//$sEndDate = date("M-d-Y", strtotime(date("M-d-Y")."-".$iCurrDayOfTheWeek." Day"));

		$sEndDate = date("M-d-Y", strtotime(date("M-d-Y")."-".($iCurrDayOfTheWeek+(7*($iSummaryWeekCount-1)))." Day"));
		
		$sEndDateYear = date("Y", strtotime(date("M-d-Y")."-".($iCurrDayOfTheWeek+(7*($iSummaryWeekCount-1)))." Day"));
		
		$iEndDateYear=intval($sEndDateYear);

		$sStartDate = date("M-d-Y", strtotime(date("M-d-Y")."-".($iCurrDayOfTheWeek+(7*$iSummaryWeekCount-1))." Day"));
		
		$sStartDateYear = date("Y", strtotime(date("M-d-Y")."-".($iCurrDayOfTheWeek+(7*$iSummaryWeekCount-1))." Day"));
		
		$iStartDateYear=intval($sStartDateYear);
		//----------------------------
		
		if ($iSummaryWeekCount==1) { //create only 1 table
			echo "<table class='summaryReportTable' id='tableSummaryReportId'>";
		}
		
		echo '<tr class="row">';

		ini_set('auto_detect_line_endings', true);

		if (!file_exists($filename)) {
			//$sDateToday = Date('Y-m-d, l');

			echo $filename." not found.";

			//echo "ExpensesTemplateTotal.csv not found.";
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
						if ($iSummaryWeekCount==1) { //add header if it's the first one;
							//if (($iColumnCount>=0) and ($iColumnCount<=$SUMMARY_COLUMN_COUNT_MAX)) {
							if (($iColumnCount>=0) and ($iColumnCount<=$SUMMARY_COLUMN_COUNT_MAX-1)) {

								//background color sky blue
								echo "<td class='tableHeaderColumn' bgcolor='#00A2E8'><b>".$cellValue."</b></td>";
							}
							else {
								echo "<td class='column'>".$cellValue."</td>";
							}
						}
					}
					else if (($iRowCount>=1) && ($iRowCount<$SUMMARY_ROW_COUNT_MAX)) {
							//echo "START: ".$sStartDate."<br/>";

							if ($iColumnCount==$SUMMARY_WEEK_COUNT_COLUMN) {
								//echo "<td class='columnCount'>".$iSummaryWeekCount."</td>";
								echo "<td class='columnWeekCount'>".($iWeekCount-$iSummaryWeekCount)."</td>";

							}
							else if ($iColumnCount==$SUMMARY_START_DATE_COLUMN) {
								echo "<td class='columnDate'>".$sStartDate."</td>";
							}
							else if ($iColumnCount==$SUMMARY_END_DATE_COLUMN) {
								echo "<td class='columnDate'>".$sEndDate."</td>";
	/*														
								echo "<td class='column'>";
									echo "<input type='text' id='cellInputId".$iWeekCount."-".$iRowCount."-".$iColumnCount."' class='inputDate' value='".$sEndDate."' min='' max=''  readonly required>";
								echo "</td>";							
	*/							
							}
							else if ($iColumnCount==$SUMMARY_TOTAL_COLUMN) {
	/*							
								echo "<td id='amtPaidId".$iRowCount."-".$iColumnCount."' class='columnAmtPaid'>".$cellValue."</td>";
	*/							
/*
								echo "<td class='columnAmtPaid'>";
									echo "<input type='text' id='cellInputId".$iWeekCount."-".$iRowCount."-".$iColumnCount."' class='inputAnswerAmtPaid' value='".$cellValue."' min='' max='' oninput='processCellInput(".$iWeekCount.",".$iRowCount.",".$iColumnCount.")'  readonly required>";
								echo "</td>";
*/								
								echo "<td class='columnAmtPaid'>";
									echo "<input type='text' id='summaryTotalColumnId".$iSummaryWeekCount."' class='inputAnswerAmtPaid' value='".$cellValue."' min='' max='' oninput=''  readonly required>";
								echo "</td>";
								
	/*
							echo "<td id='cellInputId".$iWeekCount."-".$iRowCount."-".$iColumnCount."' class='inputAnswerAmtPaid'>".$cellValue."</td>";	
	*/						
							}
							else {
								echo "<td class='column'>".$cellValue."</td>";
							}							
					}	
					
				}
				//echo '</tr><tr class="row">';
				echo '</tr>';
			  }
			  //echo '</tr>';
			  
			  fclose($handle);
			}
		}
	}
	while (($iEndDateYear!=$iPrevYear) && ($iStartDateYear!=$iPrevYear));
	
	//echo $iWeekCount;
	//$iWeekCount++;
?>  
	<tr>
		<td><b>GRAND TOTAL</b></td>
		<td></td>
		<td></td>
		<td id="summaryGrandTotalColumnId" class='columnTotal'>0.00</td>
<!--		
		<td id="summaryGrandTotalAvePerWeekColumnId" class='columnTotal'>0.00</td>
-->
		<td class='columnAmtPaid'>
			<input type='text' id='summaryGrandTotalAvePerWeekColumnId' class='inputAnswerAmtPaid' value='0.00' min='' max='' oninput=''  readonly required>
		</td>
	</tr>

	<input type="hidden" id="weekCountMaxId" value="<?php echo $iWeekCount;?>">

	</table>
	<div class="divDemoVersion">
<!--
		<span>© <b>www.usbong.ph</b> 2011~<?php echo date("Y");?>. All rights reserved.</span>
-->
		<span id="disclaimerLabelId" class="spanDisclaimerLabel">
		Demo Version. Please <a class="" target="_blank" href="https://github.com/usbong/workbook">report</a> for any errors.
		</span>
	</div>		 
  </body>
</html>