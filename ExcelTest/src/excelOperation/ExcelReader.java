package excelOperation;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import excelOperation.SheetReader.CustomTypeReader;

public class ExcelReader {

    private HashMap<String, SheetReader> mSheetReaders;

    @SuppressWarnings("resource")
    public ExcelReader(File file) throws InvalidFormatException, IOException {
        if (file == null) {
            throw new NullPointerException();
        }
        Workbook workbook;
        if (file.getPath().endsWith("xlsx")) {
            workbook = new XSSFWorkbook(file);
        } else if (file.getPath().endsWith("xls")) {
            workbook = new HSSFWorkbook(new FileInputStream(file));
        } else {
            throw new RuntimeException("file type error:" + file.getPath());
        }
        mSheetReaders = new HashMap<String, SheetReader>();
        for (int numSheet = 0; numSheet < workbook.getNumberOfSheets(); numSheet++) {
            Sheet sheet = workbook.getSheetAt(numSheet);
            if (sheet != null) {
                mSheetReaders.put(sheet.getSheetName(), new SheetReader(sheet));
            }
        }
    }

    private CustomTypeReader mCustomTypeReader = new CustomTypeReader() {
        @Override
        public CaseBean readCustomObj(String type, String val) {
            if(val == null){return null;}
            int index = (int) Double.parseDouble(val);
            return findSheetByName(type).readCase(index, mCustomTypeReader);
        }

        @Override
        public ArrayList<CaseBean> readCustomObjArray(String type, String val) {
            String[] vals = Util.toValArray(val);
            if (vals == null) { return null; }
            ArrayList<CaseBean> caseBeans = new ArrayList<CaseBean>();
            for (String valStr : vals) {
                if(valStr.matches("[0-9]*")){
                    int index = (int) Double.parseDouble(valStr);
                    caseBeans.add(findSheetByName(type).readCase(index, mCustomTypeReader));
                } else if(valStr.matches("[0-9]*[-][0-9]*")){
                    String[] valArr = valStr.split("-");
                    int index_s = (int) Double.parseDouble(valArr[0]);
                    int index_e = (int) Double.parseDouble(valArr[1]);
                    for (int i = index_s; i <= index_e; i++) {
                        caseBeans.add(findSheetByName(type).readCase(i, mCustomTypeReader));
                    }
                } else if(valStr.matches("[0-9]*[(][0-9]*[)]")){
                    String[] valArr = valStr.split("\\(");
                    int index = (int) Double.parseDouble(valArr[0]);
                    int cnt = (int) Double.parseDouble(valArr[1].substring(0, valArr[1].length() - 1));
                    for (int i = 0; i < cnt; i++) {
                        caseBeans.add(findSheetByName(type).readCase(index, mCustomTypeReader));
                    }
                } else {
                    System.out.println(valStr);
                }
            }
//                caseBeans.add(findSheetByName(type).readCase(indexs[i], mCustomTypeReader));
            return caseBeans;
        }
    };

    public void readCase(int index){
        SheetReader mainReader = mSheetReaders.get("Main");
        if(mainReader == null){return;}
        CaseBean caseBean = mainReader.readCase(index, mCustomTypeReader);
        printCaseBean(caseBean);
    }
    private void printCaseBean(CaseBean caseBean) {
        for (ValueBean valueBean : caseBean.dataList) {
            System.out.println("type:" + valueBean.type + "  key:" + valueBean.key + "  val:" + valueBean.val);
            if (valueBean.val instanceof CaseBean) {
                printCaseBean((CaseBean) valueBean.val);
            } else if (valueBean.val instanceof ArrayList) {
                ArrayList<CaseBean> caseList = (ArrayList<CaseBean>)valueBean.val;
                for (CaseBean caseBean2 : caseList) {
                    printCaseBean(caseBean2);
                }
            }
        }
        
    }
    /**
     * 查找对应名称的sheet，返回SheetReader对象
     * @param sheetName
     * @return
     */
    private SheetReader findSheetByName(String sheetName) {
        if (mSheetReaders == null) { return null; }
        return mSheetReaders.get(sheetName);
    }
}
