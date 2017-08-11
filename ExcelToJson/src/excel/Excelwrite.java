

package excel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Excelwrite {
    public static void main(String args[]) throws FileNotFoundException, IOException
    {
        XSSFWorkbook wb= new XSSFWorkbook ();
        Sheet sheet = wb.createSheet("text");
        for(int i=0; i<5 ;i++)
        {
        Row row = sheet.createRow((int) i);
        if(row.getRowNum()==0)
        {
     CellStyle style = wb.createCellStyle();//Create style
    Font font = wb.createFont();//Create font
    font.setBoldweight(Font.BOLDWEIGHT_BOLD);//Make font bold
    style.setFont(font);
   style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
style.setFillPattern(CellStyle.SOLID_FOREGROUND); 
            
            for(int ii=0;ii<5;ii++)
            {
                Cell cell =row.createCell(ii);
                cell.setCellValue("King");
                cell.setCellStyle(style);
            }
            
            i++;
            row=sheet.createRow((short)i);
        }
        for(int j=0; j<5;j++)
        {
            Cell cell = row.createCell(j);
            cell.setCellValue("none");
        }
        }
         FileOutputStream fileOut = new FileOutputStream("src/res/empty.xlsx");
         wb.write(fileOut);
        
        fileOut.close();
        System.out.println("Excel file craeted");
    }
    
}
