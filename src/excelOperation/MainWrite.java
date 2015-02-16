package excelOperation;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.omg.CORBA_2_3.portable.OutputStream;

public class MainWrite {
    private final static int START_ROW = 96;
    private final static int END_ROW = 287;
    private final static int INSERT_COL = 34;
    /**
     * @param args
     */
    public static void main(String[] args) {
        String filePath = "/home/luoyuheng/workspace/ExcelTest/file/shenqingshu.xls";
        String logPath = "/home/luoyuheng/workspace/ExcelTest/file/log.log";
        File file = new File(filePath);
        File logFile = new File(logPath);
        boolean isNewItem = true;
        try {
            BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(logFile)));
            FileInputStream is = new FileInputStream(file);
            HSSFWorkbook workbook = new HSSFWorkbook(is);
            MySheet sheet = new MySheet(workbook.getSheetAt(6));
//            System.out.println(sheet.getCell(97, 34).getStringCellValue());
//            System.out.println(sheet.getCell(97, 34).getStringCellValue());
//            os.close();
//            is.close();
            String str = null;
            int row = 0;
            while ((str = br.readLine()) != null) {
                if(str.matches("[#][0-9]*")){
                    isNewItem = true;
                    row = Integer.valueOf(str.substring(1));
                    continue;
                }
                if (isNewItem) {
                    if(str.startsWith("tests/")){
                        continue;
                    }
                    System.out.println(row+":"+str);
                    String[] strings = str.split(":");
                    String val = "Class:"+strings[0]+"\r\n"+"Line:"+strings[1] + "\r\n"+"Function:\r\n";
                    sheet.setCellValue(row-1, INSERT_COL, val);
                    isNewItem = false;
                }
            }
            System.out.println("Done!");
            FileOutputStream os  = new FileOutputStream(file);
            workbook.write(os);
            workbook.close();
            os.flush();
            os.close();
            is.close();
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

}
