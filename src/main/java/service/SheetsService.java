package service;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;
import config.SheetsConfig;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.List;

public enum SheetsService {
	INSTANCE;

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

		return originValueRange.getValues();
	}

	public UpdateValuesResponse writeDataToSheet(String destinySpreadsheetsId, String destinyRange, ValueRange destinyBody, String destinyValueInputOption) throws IOException, GeneralSecurityException {
		Sheets sheetsService = SpreadsheetService.INSTANCE.SpreadsheetService();

		return sheetsService.spreadsheets().values().update(destinySpreadsheetsId, destinyRange, destinyBody)
				.setValueInputOption(destinyValueInputOption)
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
}
