//Lembrar que as depend�ncias JAR da API devem ser adicionadas em Libraries -> Classpath, ou o programa �ao encontrar� as classes.

package exceloperations;

import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

import entities.ExcelFile;

public class ReadingExcel {

	public static List<String> listCells2 = new ArrayList<>();
	
	public static void main(String[] args) {

		try {
			Scanner sc = new Scanner(System.in);
			System.out.println("Digite o caminho do primeiro arquivo: ");
			String path1 = sc.nextLine();
			System.out.println("Digite o caminho do segundo arquivo: ");
			String path2 = sc.nextLine();
			sc.close();
			ExcelFile excelfile = new ExcelFile();
			excelfile.ReadFromFile(path1);
			excelfile.ReadFromFile(path2);
			excelfile.compareLists();
			System.out.println(listCells2);
			// System.out.println(excelfile2); // O m�todo toString funcionou para o objeto
			// excelFile mas n�o funciona para listas
			
		} catch (IllegalStateException e) {
			System.out.println("Verificar planilha: H� dados n�o string!");
		}
	}

}
