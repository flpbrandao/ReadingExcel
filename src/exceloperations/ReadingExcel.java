//Lembrar que as dependências JAR da API devem ser adicionadas em Libraries -> Classpath, ou o programa ñao encontrará as classes.

package exceloperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;

public class ReadingExcel {

	public static void main(String[] args) throws IOException {
			
		
		String excelFilePath = "C:\\temp\\workspace eclipse\\POIApache\\datafiles\\how.xlsx";
		try {
			FileInputStream inputstream = new FileInputStream(excelFilePath); // Inicializa leitor de arquivos genérico - fs
			XSSFWorkbook workbook=new XSSFWorkbook(inputstream);; // Inicializa objeto workbook (Planilha) da classe Herdada
															// XSSFWorkbook(POI Apache) a partir do arquivo fs lido
															// - Pode gerar exceção IOException se o objeto não for uma
															// planilha xlsx

			XSSFSheet sheet = workbook.getSheet("Plan1"); // Criado objeto de planilha a partir da classe XSSFSheet, a
															// partir do Workbook

			int numRows = sheet.getLastRowNum(); // Método da classe XSSFSheet que pega número de linhas da planilha
			int numCols = sheet.getRow(1).getLastCellNum(); // Método da classe XSSFSheet que pega número de colunas
															// naquela linha

			for (int r = 0; r <= numRows; r++) { // For para percorrer as linhas da planilha

				XSSFRow row = sheet.getRow(r); // Cria um objeto de linha com o método getRow na posição do for

				for (int c = 0; c < numCols; c++) { // Para cada linha, será percorrido o numero de colunas obtido
													// anteriormente

					XSSFCell cell = row.getCell(c); // Em cada coluna, será obtida a célula

					switch (cell.getCellType()) { // Verifica qual o tipo de conteúdo da célula para ler de acordo

					case STRING:
						System.out.print(cell.getStringCellValue());
						break; // Se for string, chamada o método para ler e sai do switch
					case NUMERIC:
						System.out.print(cell.getNumericCellValue());
						break;
					case BOOLEAN:
						System.out.print(cell.getBooleanCellValue());
						break;
					}
					System.out.print(" | "); // Separador para formatação sendo executado dentro do for
				}
				System.out.println(); // Separador para formatação sendo executado dentro do for das linhas
			}
			inputstream.close();
			
		}

		catch (FileNotFoundException e) {
			System.out.println("Arquivo não encontrado!");

		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
