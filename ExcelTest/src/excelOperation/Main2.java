package excelOperation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Main2 {

    private final static int START_ROW = 21;
    private final static int END_ROW = 2500;
    private final static String MAIN_PATH = "/home/luoyuheng/fujitsuEL/LINUX/";
    /**
     * @param args
     */
    public static void main(String[] args) {
        // TODO Auto-generated method stub
        String filePath = "/home/luoyuheng/workspace/ExcelTest/file/shenqingshu.xls";
        File file = new File(filePath);
        File scriptFile = new File("/home/luoyuheng/fujitsuEL/LINUX/android/grep_script.sh");
        
        try {
            scriptFile.createNewFile();
            scriptFile.setWritable(true);
            scriptFile.setReadable(true);
            scriptFile.setExecutable(true);
            InputStream is = new FileInputStream(file);
            MyFileWriter scriptFileWriter = new MyFileWriter(scriptFile);
            HSSFWorkbook workbook = new HSSFWorkbook(is);
            MySheet sheet = new MySheet(workbook.getSheetAt(6));
            scriptFileWriter.appendLine("#!/bin/sh");
            for (int i= START_ROW; i < END_ROW; i++) {
                int id = i+1;
                String subpath = sheet.getCell(i, 6).getStringCellValue();
                scriptFileWriter.appendLine("cd "+MAIN_PATH+subpath);
                scriptFileWriter.appendLine("cd ..");
                scriptFileWriter.appendLine("echo \"#"+id+"\"");
                String resId = sheet.getCell(i, 8).getStringCellValue();
                scriptFileWriter.appendLine("grep -nR \"R.string."+resId+"\" |grep java");
                scriptFileWriter.appendLine("grep -nR \"@string/"+resId+"\" |grep xml");
//                System.out.println(i+1+":"+);
            }
            scriptFileWriter.flush();
            scriptFileWriter.close();
            System.out.println("Done!!");
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

}
