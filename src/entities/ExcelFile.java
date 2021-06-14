package entities;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import exceloperations.ReadingExcel;

public class ExcelFile {

	private String path;

	public ExcelFile(String path) {
		this.path = path;
	}

	public String getPath() {
		return path;
	}

	public void setPath(String path) {
		this.path = path;
	}

	public void ReadFromFile() {
		try {
			String excelFilePath = this.getPath();
			FileInputStream inputstream = new FileInputStream(excelFilePath); // Inicializa leitor de arquivos gen�rico
																				// - fs
			XSSFWorkbook workbook = new XSSFWorkbook(inputstream);
			; // Inicializa objeto workbook (Planilha) da classe Herdada
				// XSSFWorkbook(POI Apache) a partir do arquivo fs lido
				// - Pode gerar exce��o IOException se o objeto n�o for uma
				// planilha xlsx

			XSSFSheet sheet = workbook.getSheet("Plan1"); // Criado objeto de planilha a partir da classe XSSFSheet, a
															// partir do Workbook

			int numRows = sheet.getLastRowNum(); // M�todo da classe XSSFSheet que pega n�mero de linhas da planilha
			int numCols = sheet.getRow(0).getLastCellNum(); // M�todo da classe XSSFSheet que pega n�mero de colunas
															// naquela linha

			for (int r = 0; r <= numRows; r++) { // For para percorrer as linhas da planilha

				XSSFRow row = sheet.getRow(r); // Cria um objeto de linha com o m�todo getRow na posi��o do for

				for (int c = 0; c < numCols; c++) { // Para cada linha, ser� percorrido o numero de colunas obtido
													// anteriormente

					XSSFCell cell = row.getCell(c); // Em cada linha dentro coluna, ser� obtida a c�lula

					switch (cell.getCellType()) { // Verifica qual o tipo de conte�do da c�lula para ler de acordo

					case STRING:// Se for string, chamada o m�todo para ler e sai do switch
						// System.out.print(cell.getStringCellValue());
						ReadingExcel.listCells2.add(cell);
						break;
					case NUMERIC:
						// System.out.print(cell.getNumericCellValue());
						ReadingExcel.listCells2.add(cell);
						break;
					case BOOLEAN:
						// System.out.print(cell.getBooleanCellValue());
						ReadingExcel.listCells2.add(cell);
						break;
					case BLANK:
						// System.out.print("Empty cell!");
						break;
					default:
						System.out.print("Unformatted data!");
						break;

					}
					
					// System.out.print(" | "); // Separador para formata��o sendo executado dentro
					// do for

				}
				// System.out.println(); // Separador para formata��o sendo executado dentro do
				// for das linhas
				
			}
			inputstream.close();
			workbook.close();
			

		}

		catch (FileNotFoundException e) {
			System.out.println("Arquivo n�o encontrado!");

		} catch (IOException e) {
			e.printStackTrace();
		} catch (NullPointerException e) {
			System.out.println("Erro na estrutura da planilha.");
		}


	}
	public String toString () {
		StringBuilder sb = new StringBuilder();
		for (XSSFCell cell : ReadingExcel.listCells2) {
			sb.append(cell.getStringCellValue() + " - ");
			
		}
		return sb.toString();

	}
}
