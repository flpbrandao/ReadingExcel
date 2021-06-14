//Lembrar que as dependências JAR da API devem ser adicionadas em Libraries -> Classpath, ou o programa ñao encontrará as classes.

package exceloperations;

import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;

import entities.ExcelFile;

public class ReadingExcel {
	
public static List<XSSFCell> listCells2 = new ArrayList<>();
	
	public static void main(String[] args) {
		
		
		Scanner sc = new Scanner(System.in);
		System.out.println("Digite o caminho do primeiro arquivo: ");
		String path = sc.nextLine();
		ExcelFile excelfile = new ExcelFile(path);
		excelfile.ReadFromFile();
		System.out.println("Digite o caminho do segundo arquivo: ");
		path = sc.nextLine();
		ExcelFile excelfile2 = new ExcelFile(path);
		excelfile2.ReadFromFile();
		sc.close();
		System.out.println(excelfile2); // O método toString funcionou para o objeto excelFile mas não funciona para listas.
				
	}
	public void addList (XSSFCell cell) {
		  listCells2.add(cell);
		
	}
	
}

	
	

