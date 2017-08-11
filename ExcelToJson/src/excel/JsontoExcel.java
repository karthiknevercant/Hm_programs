package excel;

import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;
import java.util.Set;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

public class JsontoExcel {
    public static String filePath = "src/res/";
    public boolean convertJsonToExcel(String fileName) throws IOException, ParseException
    {
    
          JSONParser parser = new JSONParser();
   //load Json file 
          System.out.println(filePath+fileName);
        Object obj = parser.parse(new FileReader(filePath+fileName));
   //putting object into Json Object
        JSONObject jsonObject = (JSONObject) obj;
        //System.out.println(jsonObject);

        Set kees = jsonObject.keySet();
        Iterator iterate = kees.iterator();
        //System.out.println("list");
   //For getting asset list    
        while (iterate.hasNext()) {
            String key = (String) iterate.next();

            Object list = jsonObject.get(key);

            //System.out.print("key : " + key);
            //System.out.println(" value :" + obj);

            XSSFWorkbook wb = new XSSFWorkbook();

            //System.out.println(list);

            JSONArray assetList = (JSONArray) list;

            //System.out.println("ASSET list : " + assetList);
    //For sheet list   
            for (int i = 0; i < assetList.size(); i++) {
             
                JSONObject asset = (JSONObject) assetList.get(i);

                //System.out.println(asset);
                //System.out.println();
                Set keys = asset.keySet();
                Iterator iterator = keys.iterator();
       //getting sheet one by one(text,image,audio,video) 
                while (iterator.hasNext()) {
                    
                    String kee = (String) iterator.next();
                    Object object = asset.get(kee);
                    //System.out.print("key : " + kee);
                    //System.out.println(" value :" + object);
           //crete sheet 
                    Sheet sheet = wb.createSheet(kee);

            // for header row(bold,bg color)
                    CellStyle style = wb.createCellStyle();//Create style
                    Font font = wb.createFont();//Create font
                    font.setBoldweight(Font.BOLDWEIGHT_BOLD);//Make font bold
                    style.setFont(font);
                    style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
                    style.setFillPattern(CellStyle.SOLID_FOREGROUND);

                    JSONArray arr = (JSONArray) object;

                    //System.out.println(arr);

                    //System.out.println();

                    //System.out.println("objects: ");
            //getting objects one by one(row)
                    for (int k = 0; k < arr.size(); k++) {
            //cell initial position      
                        int pos = 0;
                        //System.out.println(k);
                        JSONObject o = (JSONObject) arr.get(k);

                        Set keees = o.keySet();

                        Iterator objIterator = keees.iterator();

                        Row row = sheet.createRow(k);
            //getting object values one by one(cell)
                        while (objIterator.hasNext()) {
                            
                            String keee = (String) objIterator.next();
                            Object value = o.get(keee);
                            //System.out.print("key : " + keee);
                            //System.out.println(" value :" + value);
                            Cell cell = row.createCell(pos++);
           //Adding header (Top row)         
                            if (k == 0) {

                                cell.setCellValue(keee);
                                cell.setCellStyle(style);
           //Adding row values(cell by cell)
                            } else {
                                if (value instanceof Double) {
                                    cell.setCellValue((double) value);
                                    
                                } else if (value instanceof String) {
                                    cell.setCellValue(value.toString());
                                }

                            }

                        }
                        //System.out.println(o);
                    }
                }
            }
       //Wrting workbook into excel file     
            FileOutputStream file = new FileOutputStream("src/res/empty.xlsx");
            wb.write(file);
            file.close();
            System.out.println("Successfully converted");
        }
        return true;
    }



    public static void main(String args[]) throws IOException, ParseException {
        Scanner sc = new Scanner(System.in);
     
        
     //getting file name from user  
        System.out.println("Note : Enter File name as in res folder ");
        
        System.out.println("Enter file name along with extension json : ");
       
        String fileName = sc.next();
        JsontoExcel object = new JsontoExcel();
        boolean flag = object.convertJsonToExcel(fileName);
        if (flag) {
            System.out.println("json code created in code.json file in a res folder ");
        }
  
}
}