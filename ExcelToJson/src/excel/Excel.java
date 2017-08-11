package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Scanner;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

public class Excel {

    public static int starterRow = 1;
    public static String filePath = "src/res/";
    public static Workbook wb = null;
    public static JSONArray asset = new JSONArray();
    public static JSONArray objectList = null;
    public static JSONObject assetList = new JSONObject();
    public static ArrayList<ExcelDto> list = null;
    public static HashMap<String, ArrayList<ExcelDto>> assetListMap = new HashMap();
    public static String[] headerValues;

    public boolean convertExcelToJson(String filepath) {
        try {
         //getting Excel file
            
            FileInputStream f = new FileInputStream(filePath+filepath);
            if (filepath.endsWith(".xlsx")) {
                wb = new XSSFWorkbook(f);
            } else if (filepath.endsWith(".xls")) {
                wb = new HSSFWorkbook(f);
            } else {
                System.out.println(" File Doesn't exists");
            }
        //getting sheets one by one
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                list = new ArrayList();
                Sheet sheet = wb.getSheetAt(i);
                
                
                
               
           //Header row     
                Row headerRow = sheet.getRow(starterRow);
           //getting colunm size     
                int colSize = headerRow.getPhysicalNumberOfCells();
       //               System.out.println("Total Colunm size :" + colSize);
           //cell iterator for header row     
                Iterator<Cell> headCellIterator = headerRow.iterator();
                headerValues = new String[colSize];                  //creating memory for array based on headers
           //storing headers into array
                for (int j = 0; j < colSize; j++) {
                    Cell cell = headCellIterator.next();
                    headerValues[j] = cell.getStringCellValue();
                }
              System.out.println(sheet.getSheetName());  
              System.out.println("headers: ");
                 for (int j = 0; j < colSize; j++) {

                 System.out.println(headerValues[j]);
                 }                                              
             //row iterator    
                 Iterator<Row> iterator = sheet.iterator();
            //row iterator
               Row row = iterator.next();
            //AssetList objects (text,image,audio,video) #jsonArray
                row = iterator.next();
                objectList = new JSONArray();
             //looping row one by one
                while (iterator.hasNext()) {
        
             //row object
                    ExcelDto obj = new ExcelDto();
                    row = iterator.next();
              //cell iterator  
                    Iterator<Cell> cellIterator = row.iterator();
              //json row object    
                    JSONObject rowObject = new JSONObject();
             //looping cell one by one  
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        if(cell ==null && cell.getCellType()== Cell.CELL_TYPE_BLANK)
                            break;
                        int colIndex = cell.getColumnIndex();
                        //  System.out.println("colunm Index : " + colIndex + " ROw INdex :" + cell.getRowIndex());
                        
              //getting cell content based on type #two types( string, Numeric )          
                        if (cell.getCellTypeEnum() == CellType.STRING) {
                            //   System.out.println(" cell value : " + cell.getStringCellValue());
                            rowObject.put(headerValues[colIndex], cell.getStringCellValue());
                        } else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                            //      System.out.println("cell vallue " + cell.getNumericCellValue());
                            rowObject.put(headerValues[colIndex], cell.getNumericCellValue());
                        }
              //setting values for object(Arraylist)
                        //For Text  
                        if (wb.getSheetIndex(sheet) == 0) {
                            switch (colIndex) {
                                case 0:
                                    obj.setScreenName(cell.getStringCellValue());
                                    break;
                                case 1:
                                    obj.setId((int) cell.getNumericCellValue());
                                    break;
                                case 2:
                                    obj.setEn(cell.getStringCellValue());
                                    break;
                                case 3:
                                    obj.setAr(cell.getStringCellValue());
                                    break;
                                default:
                                    System.out.println("Error");
                            }
                        }
                        //For Image 
                        else if (wb.getSheetIndex(sheet) == 1) {
                            switch (colIndex) {
                                case 0:
                                    obj.setScreenName(cell.getStringCellValue());
                                    break;
                                case 1:
                                    obj.setId((int) cell.getNumericCellValue());
                                    break;
                                case 2:
                                    obj.setName((int) cell.getNumericCellValue());
                                    break;
                                case 3:
                                    obj.setFileName(cell.getStringCellValue());
                                    break;
                                case 4:
                                    obj.setResolution(cell.getStringCellValue());
                                    break;
                                case 5:
                                    obj.setEn(cell.getStringCellValue());
                                    break;
                                case 6:
                                    obj.setAr(cell.getStringCellValue());
                                    break;
                                default:
                                    System.out.println("Error");
                            }

                        }
                        //For Audio    
                        if (wb.getSheetIndex(sheet) == 2) {
                            switch (colIndex) {
                                case 0:
                                    obj.setScreenName(cell.getStringCellValue());
                                    break;
                                case 1:
                                    obj.setId((int) cell.getNumericCellValue());
                                    break;
                                case 2:
                                    obj.setName((int) cell.getNumericCellValue());
                                    break;
                                case 3:
                                    obj.setFileName(cell.getStringCellValue());
                                    break;
                                case 4:
                                    obj.setEn(cell.getStringCellValue());
                                    break;
                                case 5:
                                    obj.setAr(cell.getStringCellValue());
                                    break;
                                default:
                                    System.out.println("Error");
                            }
                        }
                        //For Video    
                        else if (wb.getSheetIndex(sheet) == 3) {
                            switch (colIndex) {
                                case 0:
                                    obj.setScreenName(cell.getStringCellValue());
                                    break;
                                case 1:
                                    obj.setId((int) cell.getNumericCellValue());
                                    break;
                                case 2:
                                    obj.setName((int) cell.getNumericCellValue());
                                    break;
                                case 3:
                                    obj.setFileName(cell.getStringCellValue());
                                    break;
                                case 4:
                                    obj.setResolution(cell.getStringCellValue());
                                case 5:
                                    obj.setEn(cell.getStringCellValue());
                                    break;
                                case 6:
                                    obj.setAr(cell.getStringCellValue());
                                    break;
                                default:
                                    System.out.println("Error");
                            }

                        }
                    }
                    
                        objectList.add(rowObject);        //adding object to jsonarray
                    list.add(obj);                  //adding object to  arraylist
                }
                JSONObject sheetObject = new JSONObject();
                sheetObject.put(wb.getSheetName(i), objectList);     //putting json array list into the sheet one by one(Text, Image, Audio, Video) object
                asset.add(sheetObject);                                 //adding that sheet object into Asset jsonArray

                assetListMap.put(wb.getSheetName(i), list);            //puttting arrayList into sheet hashMap
            }
            
            assetList.put("Asset List", asset);                 //putting Asset jsonArray into AssetList Object

        //Creating a empty json file  
            File file = new File("src/res/code.json");

            if (file.createNewFile()) {
                System.out.println("File is created!");
            } else {
                FileWriter writer = new FileWriter("src/res/code.json");
                writer.write("");
            }
        // writing json code into file named code     
            FileWriter fileWriter = new FileWriter("src/res/code.json");
            fileWriter.write(assetList.toJSONString());
            fileWriter.flush();
            //      System.out.println("Text object :" + assetList);
        //Exception Handling            
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
   //     System.out.println("ArrayList values : ");
        
        /*  for (ExcelDto o : list) {

                     System.out.println(o.getScreenName() + "  " + o.getId());
         }                                                                     */

        /*  System.out.println("HashMap values :");
         for (Map.Entry<String, ArrayList<ExcelDto>> entry : assetListMap.entrySet()) {
         System.out.println(entry.getKey());
         System.out.println(entry.getValue());
         System.out.println(entry.getValue().size());                        

         }                                                                */
        return true;
    }

    public static void main(String[] args) {

        Scanner sc = new Scanner(System.in);
     
        
     //getting file name from user  
        System.out.println("Note : Enter File name as in res folder ");
        
        System.out.println("Enter file name along with extension xlsx or xls : ");
       
        String fileName= sc.next();
        Excel excel = new Excel();
        boolean flag = excel.convertExcelToJson(fileName);
        if (flag) {
            System.out.println("json code created in code.json file in a res folder ");
        }

    }
}
