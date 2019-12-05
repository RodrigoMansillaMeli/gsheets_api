import com.google.api.services.sheets.v4.model.AppendValuesResponse;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetResponse;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;
import com.sun.org.apache.xerces.internal.xs.datatypes.ObjectList;
import org.apache.log4j.net.SimpleSocketServer;
import service.SheetsService;

import java.io.*;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Scanner;

public class SheetsApplication {
	public static void main(String... args) throws IOException, GeneralSecurityException {
		/**
		 * Crea una hoja vacia
		 * Podemos agregarle hojas, sino se crea por defecto con una sola
		 */
		/*String title = "Destiny Google Sheets API";
		List<Sheet> sheets = service.SheetsService.createNewSheets(4); //Probablemente se cree el google spreadsheet a mano
		String destinySpreadsheetsId = service.SheetsService.createNewSpreadsheet(sheets, title);*/
		String destinySpreadsheetsId = "1pOMVZ9AgtJsJo3-q6LcgPeHHlLr_VAE8EQKYgCeAruc"; //Usamos el archivo que creamos en ejecuciones anteriores

		/**
		 * Eliminar datos de una hoja
		 */
		String sheetNameToClear = "Hoja 1"; //Abarca toda la hoja

		SheetsService.INSTANCE.clearValuesFromSheet(destinySpreadsheetsId, sheetNameToClear);

		/**
		 * Aplicamos formato a un rango, por ahora, hardcodeado
		 */

		/*Integer destinySheetId = 704011430; //Usamos la hoja creada en ejecuciones anteriores
		Integer percentColumnNumber = 8;

		SheetsService.INSTANCE.formatPercentColumns(destinySpreadsheetsId, destinySheetId, percentColumnNumber);

		SheetsService.INSTANCE.formatHeadCells(destinySpreadsheetsId, destinySheetId);*/

		/**
		 * Leemos un SpreadSheet
		 * Con el spreadsheetId de ejemplo que tenemos en el Drive
		 */
		final String originSpreadsheetsId = "1ZHoIuRlTTRNN4jUw7OFAEu04XgeAMAIjPR3ti7otA2U"; //Tiene los datos de origen
		final String originHeadersRange = "Datos!A1:ZZZ1";
		final String originDataRange = "Datos!A2:ZZZ"; //Seteamos nombre de la hoja, donde arrancan los datos y hasta cual columna (COPIA TODO)
		//final String originRange = "Datos"; //Seteamos para que copie toda la hoja

		List<List<Object>> originHeadersValues = SheetsService.INSTANCE.retrieveDataFromSheet(originSpreadsheetsId, originHeadersRange);
		List<List<Object>> originDataValues = SheetsService.INSTANCE.retrieveDataFromSheet(originSpreadsheetsId, originDataRange);

		//Probamos la lectura con un print.
		/*if (originValues == null || originValues.isEmpty()) {
			System.out.println("No data found.");
		} else {
			System.out.println("Name, Major, Class Level, Numeric Column, Float Column");
			for (List row : originValues) {
				// Print columns A and E, which correspond to indices 0 and 4
				System.out.printf("%s, %s, %s, %s, %s\n", row.get(0), row.get(4), row.get(2), row.get(6), row.get(7));
			}
		}*/

		/**
		 * Con los valores de la lectura anterior, actualizamos el nuevo
		 * Spreadsheet creado con estos datos.
		 * Solo la hoja que necesitamos
		 *
		 * ESTO NO BORRA LA HOJA, SOLO PISA EL RANGO QUE SE ACTUALIZA
		 */
		String destinyDataRange = "Hoja 1!A2:ZZZ";
		String destinyHeadersRange = "Hoja 1!A1:ZZZ1";
		String destinyValueInputOption = "USER_ENTERED";

		ValueRange destinyHeadersBody = new ValueRange()
				.setValues(originHeadersValues);

		ValueRange destinyDataBody = new ValueRange()
				.setValues(originDataValues);

		//Escribe los datos sobre sheet destiny.

		System.out.println(destinyValueInputOption);
		SheetsService.INSTANCE.writeDataToSheet(destinySpreadsheetsId, destinyHeadersRange, destinyHeadersBody, destinyValueInputOption);
		SheetsService.INSTANCE.writeDataToSheet(destinySpreadsheetsId, destinyDataRange, destinyDataBody, destinyValueInputOption);

		/**
		 * Lee datos desde CSV
		 */

		String fileName= "/Users/abasile/eclipse-workspace/testeo/src/testeo/datos.csv";
		File file= new File(fileName);

		// this gives you a 2-dimensional array of strings
		List<List<Object>> lines = new ArrayList<>();
		Scanner inputStream;

		System.out.println(lines);

		try{
			inputStream = new Scanner(file);

			while(inputStream.hasNext()){
				String line= inputStream.next();
				String[] values = line.split(",");

				// this adds the currently parsed line to the 2-dimensional string array
				lines.add(Arrays.<Object>asList(values));


			}

			inputStream.close();
		}catch (FileNotFoundException e) {
			e.printStackTrace();
		}

		//System.out.println(lines);



		/**
		 * Agregar registros nuevos en la hoja
		 */
		String appendToSheet = "Hoja 1";
		destinyValueInputOption = "USER_ENTERED";
		String insertDataOption = "INSERT_ROWS";
		ValueRange appendBody = new ValueRange()
				.setValues(lines);

		AppendValuesResponse appendResult = SheetsService.INSTANCE.appendDataToSheet(destinySpreadsheetsId, appendToSheet, appendBody, destinyValueInputOption, insertDataOption);

		System.out.printf("\n%s range updated.\n", appendResult.getTableRange());


		/**
		 * Actualizar a partir de un rango!
		 */
		/*String updateInRange = "G1";
		ValueRange updateBody = new ValueRange()
				.setValues(Arrays.asList(
						Arrays.<Object>asList("TestColumn", "TESTCOLUMN")
				));

		UpdateValuesResponse updateResult = SheetsService.INSTANCE.updateDataToSheet(destinySpreadsheetsId, updateInRange, updateBody, destinyValueInputOption);*/

		/**
		 * Lectura de csv
		 */



	}
}
