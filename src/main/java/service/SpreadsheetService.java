package service;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.Sheet;
import config.SheetsConfig;

import java.io.IOException;
import java.security.GeneralSecurityException;

public enum SpreadsheetService {

	INSTANCE;

	private final String APPLICATION_NAME = "Google Sheets API Java Quickstart";
	private final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();

	Sheets SpreadsheetService() throws IOException, GeneralSecurityException {
		final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
		Credential credential = SheetsConfig.getCredentials(HTTP_TRANSPORT, JSON_FACTORY);
		return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
				.setApplicationName(APPLICATION_NAME)
				.build();
	}
}
