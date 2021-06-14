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
			FileInputStream inputstream = new FileInputStream(excelFilePath); // Inicializa leitor de arquivos genérico
																				// - fs
			XSSFWorkbook workbook = new XSSFWorkbook(inputstream);
			; // Inicializa objeto workbook (Planilha) da classe Herdada
				// XSSFWorkbook(POI Apache) a partir do arquivo fs lido
				// - Pode gerar exceção IOException se o objeto não for uma
				// planilha xlsx

			XSSFSheet sheet = workbook.getSheet("Plan1"); // Criado objeto de planilha a partir da classe XSSFSheet, a
															// partir do Workbook

			int numRows = sheet.getLastRowNum(); // Método da classe XSSFSheet que pega número de linhas da planilha
			int numCols = sheet.getRow(0).getLastCellNum(); // Método da classe XSSFSheet que pega número de colunas
															// naquela linha

			for (int r = 0; r <= numRows; r++) { // For para percorrer as linhas da planilha

				XSSFRow row = sheet.getRow(r); // Cria um objeto de linha com o método getRow na posição do for

				for (int c = 0; c < numCols; c++) { // Para cada linha, será percorrido o numero de colunas obtido
													// anteriormente

					XSSFCell cell = row.getCell(c); // Em cada linha dentro coluna, será obtida a célula

					switch (cell.getCellType()) { // Verifica qual o tipo de conteúdo da célula para ler de acordo

					case STRING:// Se for string, chamada o método para ler e sai do switch
						ReadingExcel.listCells2.add(cell);
						break;
					case NUMERIC:
						ReadingExcel.listCells2.add(cell);
						break;
					case BOOLEAN:
						ReadingExcel.listCells2.add(cell);
						break;
					case BLANK:
						break;
					default:
						System.out.print("Unformatted data!");
						break;

					}
					
					

				}
				
				
			}
			inputstream.close();
			workbook.close();
			

		}

		catch (FileNotFoundException e) {
			System.out.println("Arquivo não encontrado!");

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
