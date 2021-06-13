//Lembrar que as depend�ncias JAR da API devem ser adicionadas em Libraries -> Classpath, ou o programa �ao encontrar� as classes.

package exceloperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;

public class ReadingExcel {

	public static void main(String[] args) throws IOException {
			
		
		String excelFilePath = "C:\\temp\\workspace eclipse\\POIApache\\datafiles\\how.xlsx";
		try {
			FileInputStream inputstream = new FileInputStream(excelFilePath); // Inicializa leitor de arquivos gen�rico - fs
			XSSFWorkbook workbook=new XSSFWorkbook(inputstream);; // Inicializa objeto workbook (Planilha) da classe Herdada
															// XSSFWorkbook(POI Apache) a partir do arquivo fs lido
															// - Pode gerar exce��o IOException se o objeto n�o for uma
															// planilha xlsx

			XSSFSheet sheet = workbook.getSheet("Plan1"); // Criado objeto de planilha a partir da classe XSSFSheet, a
															// partir do Workbook

			int numRows = sheet.getLastRowNum(); // M�todo da classe XSSFSheet que pega n�mero de linhas da planilha
			int numCols = sheet.getRow(1).getLastCellNum(); // M�todo da classe XSSFSheet que pega n�mero de colunas
															// naquela linha

			for (int r = 0; r <= numRows; r++) { // For para percorrer as linhas da planilha

				XSSFRow row = sheet.getRow(r); // Cria um objeto de linha com o m�todo getRow na posi��o do for

				for (int c = 0; c < numCols; c++) { // Para cada linha, ser� percorrido o numero de colunas obtido
													// anteriormente

					XSSFCell cell = row.getCell(c); // Em cada coluna, ser� obtida a c�lula

					switch (cell.getCellType()) { // Verifica qual o tipo de conte�do da c�lula para ler de acordo

					case STRING:
						System.out.print(cell.getStringCellValue());
						break; // Se for string, chamada o m�todo para ler e sai do switch
					case NUMERIC:
						System.out.print(cell.getNumericCellValue());
						break;
					case BOOLEAN:
						System.out.print(cell.getBooleanCellValue());
						break;
					}
					System.out.print(" | "); // Separador para formata��o sendo executado dentro do for
				}
				System.out.println(); // Separador para formata��o sendo executado dentro do for das linhas
			}
			inputstream.close();
			
		}

		catch (FileNotFoundException e) {
			System.out.println("Arquivo n�o encontrado!");

		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
