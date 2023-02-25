package apachePoi;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class LeituraDeExcelComApachePoi {
	public static void main(String[] args) throws IOException {
		// Cria um objeto FileInputStream para ler o arquivo
		FileInputStream entrada = new FileInputStream(
				"C:\\workspace-java\\AulasJavaAvancadas\\src\\trabalhandoComApachePoi\\arquivo_excel.xls");

		// Cria um objeto HSSFWorkbook a partir do arquivo
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(entrada);

		// Obtém a primeira planilha do arquivo
		HSSFSheet planilhaHssfSheet = hssfWorkbook.getSheetAt(0);

		// Cria um Iterator para percorrer as linhas da planilha
		Iterator<Row> linhasIterator = planilhaHssfSheet.iterator();

		// Cria uma lista vazia para armazenar objetos ArquivoPessoas
		List<ArquivoPessoas> pessoas = new ArrayList<ArquivoPessoas>();

		// Enquanto houver linhas na planilha
		while (linhasIterator.hasNext()) {
			// Obtém a próxima linha e a atribui à variável 'row'
			Row row = (Row) linhasIterator.next();

			// Cria um Iterator para percorrer as células da linha atual
			Iterator<Cell> cellulaIterator = row.iterator();

			// Cria um objeto ArquivoPessoas para armazenar os dados da linha atual
			ArquivoPessoas pessoa = new ArquivoPessoas();

			// Enquanto houver células na linha atual
			while (cellulaIterator.hasNext()) {
				// Obtém a próxima célula e a atribui à variável 'cell'
				Cell cell = (Cell) cellulaIterator.next();

				// Verifica a posição da célula na linha e armazena o valor na propriedade correspondente do objeto ArquivoPessoas
				switch (cell.getColumnIndex()) {
				case 0:
					pessoa.setNomeString(cell.getStringCellValue());
					break;
				case 1:
					pessoa.setEmailString(cell.getStringCellValue());
					break;
				case 2:
					pessoa.setDataDeNascimentoString(cell.getStringCellValue());
					break;
				case 3:
					// Converte o valor da célula para um número inteiro e armazena na propriedade correspondente do objeto ArquivoPessoas
					pessoa.setIdade(Double.valueOf(cell.getNumericCellValue()).intValue());
					break;
				}
			}

			// Adiciona o objeto ArquivoPessoas à lista de pessoas
			pessoas.add(pessoa);
		}

		// Fecha o objeto FileInputStream
		entrada.close();

		// Imprime os objetos ArquivoPessoas na lista de pessoas
		for (ArquivoPessoas arquivoPessoas : pessoas) {
			System.out.println(pessoas);
		}
	}
}
