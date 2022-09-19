package adquira;

import java.awt.Color;
import java.awt.Font;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
  
public class TestePoi {
        
       private static final String fileName = "C:/Viegas/desenvolvimento/Selenium/Adquira/relatorios/teste.xlsx";
       private static List<Pedido> listaPedidos = new ArrayList<Pedido>();
  
       @SuppressWarnings("deprecation")
	public static void main(String[] args) throws IOException {
  
             //HSSFWorkbook workbook = new HSSFWorkbook();
             XSSFWorkbook workbook = new XSSFWorkbook();
             XSSFSheet sheetPedidos = workbook.createSheet("Pedidos");
             
             sheetPedidos.setColumnWidth(0, "Contract Number".length() * 260);
             sheetPedidos.setColumnWidth(1, "Número Pedido".length() * 270);
             sheetPedidos.setColumnWidth(2, "Data Pedido".length() * 250);
             sheetPedidos.setColumnWidth(3, "Valor Pedido".length() * 250);
               

             Pedido pedido = new Pedido();
             pedido.setNumeroContractNumber("123");
             //pedido.setData("10/01/2009");
             pedido.setNumero("777");
             //pedido.setValor("R$ 1000,00");
             
             listaPedidos.add(pedido);
             
             Pedido pedido2 = new Pedido();
             pedido2.setNumeroContractNumber("456");
             //pedido2.setData("10/06/1937");
             pedido2.setNumero("1");
             //pedido2.setValor("R$ 1000000000,00");
             
             listaPedidos.add(pedido2);
             
             //XSSFCellStyle style = workbook.createCellStyle();
             //HSSFPalette palette = ((XSSFWorkbook) workbook).getCustomPalette();
             
             // Configuração do Style da célula
             // Negrito
             XSSFCellStyle styleNegrito = workbook.createCellStyle();
             //XSSFColor myColor = new XSSFColor(Color.LIGHT_GRAY);
             XSSFFont font = workbook.createFont();
             //style.setFillForegroundColor(myColor);
             //style.setFillBackgroundColor(myColor);
             //styleNegrito.setFillPattern(FillPatternType.SOLID_FOREGROUND);
             font.setBold(true);
             styleNegrito.setFont(font);
               
             int rownum = 0;
             int cellnumero = 0;
             // Cabeçalho
             Row row = sheetPedidos.createRow((short)rownum++);
             sheetPedidos.autoSizeColumn((short) rownum++);
             Cell cellCabecalhoContractNumber = row.createCell(cellnumero++);
             cellCabecalhoContractNumber.setCellValue("Contract Number");

             Cell cellCabecalhoNumero = row.createCell(cellnumero++);
             cellCabecalhoNumero.setCellValue("Número Pedido");
             
             Cell cellCabecalhoData = row.createCell(cellnumero++);
             cellCabecalhoData.setCellValue("Data Pedido");
             
             Cell cellCabecalhoValor = row.createCell(cellnumero++);
             cellCabecalhoValor.setCellValue("Valor Pedido");
             cellCabecalhoValor.setCellStyle(styleNegrito); // this command seems to fail
             
             //
             //adicionando bordas
             //CellStyle style = workbook.createCellStyle(); //criei o CellStyle
/*             style.setAlignment(HSSFCellStyle.ALIGN_CENTER); //texto centralizado
             style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND); //não sei kk
             style.setFillForegroundColor(HSSFColor.LEMON_CHIFFON.index); //cor de fundo
             style.setBorderBottom(CellStyle.BORDER_THIN); //borda de baixo
             style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); //cor borda baixo
             style.setBorderLeft(CellStyle.BORDER_THIN); //borda da esquerda
             style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); //cor borda esquerda
             style.setBorderRight(CellStyle.BORDER_THIN); //borda direita
             style.setRightBorderColor(IndexedColors.BLACK.getIndex()); //cor borda direita
             style.setBorderTop(CellStyle.BORDER_THIN); //borda de cima
             style.setTopBorderColor(IndexedColors.BLACK.getIndex()); //cor borda cima
*/             
             
             

             
             
             //style.setFillPattern(CellStyle.SOLID_FOREGROUND);
             //
             
             
             for (Pedido pedidos : listaPedidos) {
                 row = sheetPedidos.createRow(rownum++);
                 int cellnum = 0;
                 
                 Cell cellContractNumber = row.createCell(cellnum++);
                 cellContractNumber.setCellValue(pedidos.getNumeroContractNumber());
                 
                 Cell cellNumero = row.createCell(cellnum++);
                 cellNumero.setCellValue(pedidos.getNumero());

                 Cell cellData = row.createCell(cellnum++);
                 cellData.setCellValue("18/03/2019");
                 
                 Cell cellValor = row.createCell(cellnum++);
                 cellValor.setCellValue(pedidos.getValor());
             }
               
             try {
                 FileOutputStream out = 
                         new FileOutputStream(new File(TestePoi.fileName));
                 workbook.write(out);
                 out.close();
                 System.out.println("Arquivo Excel criado com sucesso!");
                   
             } catch (FileNotFoundException e) {
                 e.printStackTrace();
                    System.out.println("Arquivo não encontrado!");
             } catch (IOException e) {
                 e.printStackTrace();
                    System.out.println("Erro na edição do arquivo!");
             }
  
       }
  
}