package apachePoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class EditandoExcelComApachePoi {
    public static void main(String[] args) throws IOException {
        // Criando um objeto do tipo "File" que representa o arquivo Excel
        File file = new File("C:\\workspace-java\\ApachePoi\\src\\apachePoi\\arquivo_excel.xls");
        
        // Criando um objeto do tipo "FileInputStream" para ler o arquivo Excel
        FileInputStream entrada = new FileInputStream(file);
        
        // Criando um objeto do tipo "HSSFWorkbook" que representa o arquivo Excel
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(entrada);
        
        // Obtendo a primeira planilha do arquivo Excel
        HSSFSheet planilha = hssfWorkbook.getSheetAt(0);
        
        // Obtendo um objeto "Iterator" para percorrer as linhas da planilha
        Iterator<Row> linhaIterator = planilha.iterator();
        
        // Percorrendo todas as linhas da planilha
        while(linhaIterator.hasNext()){
            // Obtendo a próxima linha da planilha
            Row linha = linhaIterator.next();
            
            // Obtendo o valor da primeira célula da linha (coluna 0)
            String ValorCelular = linha.getCell(0).getStringCellValue();
            
            // Modificando o valor da primeira célula da linha
            linha.getCell(0).setCellValue(ValorCelular + "* valor gravado na aula");
        }
        
        // Fechando o objeto "FileInputStream"
        entrada.close();
        
        // Criando um objeto do tipo "FileOutputStream" para escrever no arquivo Excel
        FileOutputStream saida = new FileOutputStream(file);
        
        // Escrevendo as alterações no arquivo Excel
        hssfWorkbook.write(saida);
        
        // Liberando os recursos do objeto "FileOutputStream"
        saida.flush();
        saida.close();
        
        // Imprimindo uma mensagem de sucesso
        System.out.println("Planilha atualizada");
        
        // fechando o objeto "hssfWorkbook"
        hssfWorkbook.close();
    }
}
