//Lembrar que as dependências JAR da API devem ser adicionadas em Libraries -> Classpath, ou o programa ñao encontrará as classes.

package exceloperations;

import java.util.Scanner;

import entities.ExcelFile;

public class ReadingExcel {

	public static void main(String[] args) {

		Scanner sc = new Scanner(System.in);

		System.out.println("Digite o caminho do arquivo: ");
		String path = sc.nextLine();

		sc.close();

		ExcelFile excelfile = new ExcelFile(path);
		excelfile.ReadFromFile();

	}
}
