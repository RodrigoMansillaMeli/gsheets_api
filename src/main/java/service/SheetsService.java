package service;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;
import org.apache.log4j.Logger;
import org.w3c.dom.ranges.Range;
import sun.awt.SunHints;

import javax.xml.soap.Text;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.List;

public enum SheetsService {

	INSTANCE;

	private static final Logger logger = Logger.getLogger(SheetsService.class);

	public List<Sheet> createNewSheets(int numOfSheets) {
		List<Sheet> sheets = new ArrayList<Sheet>();
		for (int i = 0; i < numOfSheets; i++) {
			String sheetName = "Hoja " + (i + 1);
			sheets.add(new Sheet().setProperties(new SheetProperties().setTitle(sheetName)));
		}
		return sheets;
	}

	public String createNewSpreadsheet(List<Sheet> sheets, String title) throws IOException, GeneralSecurityException {
		Sheets sheetsService = SpreadsheetService.INSTANCE.SpreadsheetService();

		Spreadsheet spreadsheet = new Spreadsheet()
				.setProperties(new SpreadsheetProperties()
						.setTitle(title))
				.setSheets(sheets);

		spreadsheet = sheetsService.spreadsheets().create(spreadsheet)
				.setFields("spreadsheetId")
				.execute();

		return spreadsheet.getSpreadsheetId();
	}

	public void clearValuesFromSheet(String destinySpreadsheetsId, String sheetNameToClear) throws IOException, GeneralSecurityException {
		Sheets sheetsService = SpreadsheetService.INSTANCE.SpreadsheetService();

		sheetsService.spreadsheets().values().clear(destinySpreadsheetsId, sheetNameToClear, new ClearValuesRequest()).execute();
	}

	public List<List<Object>> retrieveDataFromSheet(String originSpreadsheetsId, String originRange) throws IOException, GeneralSecurityException {
		Sheets sheetsService = SpreadsheetService.INSTANCE.SpreadsheetService();

		ValueRange originValueRange = sheetsService.spreadsheets().values()
				.get(originSpreadsheetsId, originRange)
				.execute();

		Sheets.Spreadsheets.Get request = sheetsService.spreadsheets().get(originSpreadsheetsId);

		return originValueRange.getValues();
	}

	public UpdateValuesResponse writeDataToSheet(String destinySpreadsheetsId, String destinyRange, ValueRange destinyBody, String destinyValueInputOption) throws IOException, GeneralSecurityException {
		Sheets sheetsService = SpreadsheetService.INSTANCE.SpreadsheetService();

		return sheetsService.spreadsheets().values().update(destinySpreadsheetsId, destinyRange, destinyBody)
				.setValueInputOption(destinyValueInputOption)
				.setResponseValueRenderOption("FORMATTED_VALUE")
				.execute();
	}

	public AppendValuesResponse appendDataToSheet(String destinySpreadsheetsId, String appendToSheet, ValueRange appendBody, String destinyValueInputOption, String insertDataOption) throws IOException, GeneralSecurityException {
		Sheets sheetsService = SpreadsheetService.INSTANCE.SpreadsheetService();

		return sheetsService.spreadsheets().values()
				.append(destinySpreadsheetsId, appendToSheet, appendBody)
				.setValueInputOption(destinyValueInputOption)
				.setInsertDataOption(insertDataOption)
				.setIncludeValuesInResponse(true)
				.execute();
	}

	public UpdateValuesResponse updateDataToSheet(String destinySpreadsheetsId, String updateInRange, ValueRange updateBody, String destinyValueInputOption) throws IOException, GeneralSecurityException {
		Sheets sheetsService = SpreadsheetService.INSTANCE.SpreadsheetService();

		return sheetsService.spreadsheets().values()
				.update(destinySpreadsheetsId, updateInRange, updateBody)
				.setValueInputOption("RAW")
				.execute();
	}

	public void formatHeadCells(String destinySpreadsheetsId, Integer destinySheetId) throws IOException, GeneralSecurityException {
		Sheets sheetsService = SpreadsheetService.INSTANCE.SpreadsheetService();

		/**
		 * Color Structure
		 * Visit https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/other#Color
		 **/
		Color myBackgroundColor = new Color();
		myBackgroundColor.setRed(0.71F);
		myBackgroundColor.setGreen(0.85F);
		myBackgroundColor.setBlue(0.29F);
		myBackgroundColor.setAlpha(1F);

		Color myForegroundColor = new Color();
		myForegroundColor.setRed(0.17F);
		myForegroundColor.setGreen(0.19F);
		myForegroundColor.setBlue(0.44F);
		myForegroundColor.setAlpha(1F);

		Color myLineBorderColor = new Color();
		myLineBorderColor.setRed(1F);
		myLineBorderColor.setGreen(1F);
		myLineBorderColor.setBlue(1F);
		myLineBorderColor.setAlpha(1F);

		/**
		 * Number Format
		 * Visit https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells#NumberFormat
		 **/

		NumberFormat myTextNumber = new NumberFormat().setType("TEXT");
		NumberFormat myNumberNumber = new NumberFormat().setType("NUMBER");
		NumberFormat myPercentNumber = new NumberFormat().setType("PERCENT");

		/**
		 * Text Format
		 * Visit https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/other#TextFormat
		 **/

		TextFormat myText = new TextFormat();
		myText.setForegroundColor(myForegroundColor);
		myText.setFontFamily("Hammersmith One"); //Visit https://fonts.google.com/
		myText.setFontSize(10);
		myText.setBold(true);
		myText.setItalic(false);
		myText.setStrikethrough(false);
		myText.setUnderline(false);

		/**
		 * Borders Format
		 * Visit https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells#Borders
		 **/
		Border lineBorder = new Border();
		lineBorder.setStyle("SOLID"); // Visit https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells#Style
		lineBorder.setWidth(10); // Deprecated
		lineBorder.setColor(myLineBorderColor);

		Borders titleBorders = new Borders();
		titleBorders.setTop(lineBorder); //As example, we assing the same border over all sides
		titleBorders.setBottom(lineBorder);
		titleBorders.setLeft(lineBorder);
		titleBorders.setRight(lineBorder);

		/**
		 * Cell Format
		 * Visit https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells
		 **/
		CellFormat myFormat = new CellFormat();
		myFormat.setHorizontalAlignment("CENTER"); // TYPE_ENUM : HORIZONTAL_ALIGN_UNSPECIFIED - LEFT - CENTER - RIGHT
		myFormat.setVerticalAlignment("MIDDLE"); // TYPE_ENUM : VERTICAL_ALIGN_UNSPECIFIED - BOTTOM - MIDDLE - TOP
		myFormat.setNumberFormat(myTextNumber);
		myFormat.setTextFormat(myText);
		myFormat.setBackgroundColor(myBackgroundColor); // red&blue background
		myFormat.setBorders(titleBorders);

		/**
		 * Cell Format
		 * Visit https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells#celldata
		 **/
		CellData cellData = new CellData();
		cellData.setUserEnteredFormat(myFormat);

		/**
		 * RepeatCellRequest
		 * Visit https://developers.google.com/resources/api-libraries/documentation/sheets/v4/java/latest/com/google/api/services/sheets/v4/model/RepeatCellRequest.html
		 * Appling cell format to a range
		 **/
		GridRange gridRange = new GridRange();
		gridRange.setSheetId(destinySheetId);
		gridRange.setStartRowIndex(0);
		gridRange.setEndRowIndex(1);
		gridRange.setStartColumnIndex(0);
		gridRange.setEndColumnIndex(8);

		RepeatCellRequest repeatCellRequest = new RepeatCellRequest();
		repeatCellRequest.setFields("*");
		repeatCellRequest.setRange(gridRange);
		repeatCellRequest.setCell(cellData);

		/**
		 * Requests list
		 */
		List<Request> requests = new ArrayList<Request>();
		requests.add(new Request().setRepeatCell(repeatCellRequest));

		BatchUpdateSpreadsheetRequest body =
				new BatchUpdateSpreadsheetRequest()
						.setIncludeSpreadsheetInResponse(true)
						.setResponseIncludeGridData(true)
						.setRequests(requests);

		sheetsService.spreadsheets().batchUpdate(destinySpreadsheetsId, body).setFields("*").execute();
	}

	public void formatPercentColumns(String destinySpreadsheetsId, Integer destinySheetId, Integer columnNumber) throws IOException, GeneralSecurityException {
		Sheets sheetsService = SpreadsheetService.INSTANCE.SpreadsheetService();

		/**
		 * Number Format
		 * Visit https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells#NumberFormat
		 **/
		NumberFormat myPercentNumber = new NumberFormat().setType("PERCENT").setPattern("0.##%");

		/**
		 * Cell Format
		 * Visit https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells
		 **/
		CellFormat myFormat = new CellFormat();

		myFormat.setNumberFormat(myPercentNumber);

		/**
		 * Cell Data
		 * Visit https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells#celldata
		 **/
		CellData cellData = new CellData();
		cellData.setUserEnteredValue(new ExtendedValue().setNumberValue(0.211));
		cellData.setUserEnteredFormat(myFormat);

		/**
		 * RepeatCellRequest
		 * Visit https://developers.google.com/resources/api-libraries/documentation/sheets/v4/java/latest/com/google/api/services/sheets/v4/model/RepeatCellRequest.html
		 * Appling cell format to a range
		 **/
		GridRange gridRange = new GridRange();
		gridRange.setSheetId(destinySheetId);
		gridRange.setStartRowIndex(1);
		gridRange.setStartColumnIndex(columnNumber-1);
		gridRange.setEndColumnIndex(columnNumber);



		MergeCellsRequest mergeCellsRequest = new MergeCellsRequest();
		mergeCellsRequest.setRange(gridRange);
		mergeCellsRequest.setMergeType("MERGE_COLUMNS");



		RepeatCellRequest repeatCellRequest = new RepeatCellRequest();
		repeatCellRequest.setFields("*");
		repeatCellRequest.setRange(new GridRange()
				.setSheetId(destinySheetId)
				.setStartRowIndex(1)
				.setStartColumnIndex(columnNumber-1)
				.setEndColumnIndex(columnNumber));
		repeatCellRequest.setCell(cellData);


		/**
		 * Requests list
		 */
		List<Request> requests = new ArrayList<Request>();
		requests.add(new Request().setMergeCells(mergeCellsRequest));
		requests.add(new Request().setRepeatCell(repeatCellRequest));

		BatchUpdateSpreadsheetRequest body =
				new BatchUpdateSpreadsheetRequest()
						.setIncludeSpreadsheetInResponse(true)
						.setResponseIncludeGridData(true)
						.setRequests(requests);

		sheetsService.spreadsheets().batchUpdate(destinySpreadsheetsId, body).setFields("*").execute();
	}
}
