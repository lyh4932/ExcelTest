package excelOperation;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.ObjectInputStream.GetField;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.PaneInformation;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.AutoFilter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellRange;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

public class Main {
    /**
     * @param args
     * @throws IOException
     */
    public static void main(String[] args){
        String filePath = "/home/luoyuheng/workspace/ExcelTest/file";
        File dir = new File(filePath);
        if (dir.isDirectory()) {
            File file[] = dir.listFiles();
            for (int i = 0; i < file.length; i++) {
//                System.out.println(file[i].getPath());
                if (!Util.equals(file[i].getPath(), "/home/luoyuheng/workspace/ExcelTest/file/file1.xlsx")) {
                    continue;
                }
                try {
                    ExcelReader excelReader = new ExcelReader(file[i]);
                    excelReader.readCase(1);
                } catch (InvalidFormatException | IOException e) {
                    e.printStackTrace();
                }
//                InputStream is = new FileInputStream(file[i]);
//                XSSFWorkbook workbook = new XSSFWorkbook(is);
//                for (int numSheet = 0; numSheet < workbook.getNumberOfSheets(); numSheet++) {
//                    XSSFSheet sheet = workbook.getSheetAt(0);
//                    if (sheet == null) {
//                        continue;
//                    }
//                    System.out.println(sheet.getSheetName());
////                    System.out.println("last:"+sheet.getLastRowNum());
//                    SheetReader reader = new SheetReader(sheet);
//                    System.out.println(reader.caseToJson(1).toString());
//                    reader.readCase(2);
////                    System.out.println(SheetUtil.getCell(sheet, 1, 3).getCellType());
////                    for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
////                        XSSFRow hssfRow = sheet.getRow(rowNum);
////                        if (hssfRow != null) {
////                            for (int columeNum = 0; columeNum < hssfRow.getLastCellNum(); columeNum++) {
////                                XSSFCell no = hssfRow.getCell(columeNum);
////                                if (no != null) {
////                                    if(XSSFCell.CELL_TYPE_STRING == no.getCellType()){
////                                        System.out.println(rowNum+":"+no.getCTCell().getR()+":"+no.getStringCellValue());
////                                    } else {
////                                        System.out.println(rowNum+":"+no.getCTCell().getR()+":"+no.getCellType());
////                                    }
////                                }
////                            }
////                        }
////                    }
//                }
            }
        }
    }

}
