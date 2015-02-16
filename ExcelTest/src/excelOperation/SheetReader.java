package excelOperation;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

/**
 * @author luoyuheng
 *
 */
public class SheetReader {
    private MySheet mSheet;
    private int mColTypeIndex;
    private int mColKeyIndex;
    private int mColValIndex;
    private int mLastRowNum;
    private int mBeginRowNum;
    private int mCaseNum;
    private int mColFirstCaseIndex;

    public SheetReader(Sheet sheet) {
        mSheet = new MySheet(sheet);
        init();
    }

    private void init() {
        if (mSheet.getSheet() == null) {
            return;
        }
        mLastRowNum = mSheet.getSheet().getLastRowNum();
        for (int rowNum = 0; rowNum <= mLastRowNum; rowNum++) {
            Row row = (Row) mSheet.getSheet().getRow(rowNum);

            if (row == null) {
                continue;
            }
            for (int columeNum = 0; columeNum < row.getLastCellNum(); columeNum++) {
               Cell cell = row.getCell(columeNum);
                if (cell != null && cell.getCellType() == Cell.CELL_TYPE_STRING) {
                    String cellStr = cell.getStringCellValue();
                    if (!cellStr.startsWith("#")) {
                        continue;
                    }
                    switch (cellStr.substring(1)) {
                    case "type":
                        mColTypeIndex = cell.getColumnIndex();
                        break;
                    case "key":
                        mColKeyIndex = cell.getColumnIndex();
                        break;
                    case "val":
                        mColValIndex = cell.getColumnIndex();
                        break;
                    case "case":
                        CellRangeAddress cellRange = mSheet.getMergeRegionIndex(cell);
                        if (cellRange != null) {
                            mColFirstCaseIndex = cellRange.getFirstColumn();
                            mCaseNum = cellRange.getLastColumn() - cellRange.getFirstColumn() + 1;
                        } else {
                            mColFirstCaseIndex = cell.getColumnIndex();
                            mCaseNum = 1;
                        }
                        mBeginRowNum = cell.getRowIndex() + 2;
                    default:
                        break;
                    }
                }
            }
        }
    }

    /**
     * 读取当前sheet中某一case数据，返回CaseBean对象
     * @param index
     * @param customTypeReader
     * @return
     */
    public CaseBean readCase(int index, CustomTypeReader customTypeReader) {
        if (index > mCaseNum) { return null; }
        int colCase = mColFirstCaseIndex + index - 1;
//        String caseName = "case" + (int) mSheet.getCell(mBeginRowNum - 1, colCase).getNumericCellValue();
        CaseBean caseBean = new CaseBean();
        for (int row = mBeginRowNum; row <= mLastRowNum; row++) {
            if (Util.equals(mSheet.getCellStringValue(row, colCase), "○")) {
                String type = mSheet.getCellStringValue(row, mColTypeIndex);
                String key = mSheet.getCellStringValue(row, mColKeyIndex);
                Object val = null;
                if (Util.equals(type, "String")) {
                    val = mSheet.getCellStringValue(row, mColValIndex);
                } else if (Util.equals(type, "int")) {
                    val = mSheet.getCellIntValue(row, mColValIndex);
                } else if (Util.equals(type, "boolean")) {
                    val = mSheet.getCellBooleanValue(row, mColValIndex);
                } else if (Util.equals(type, "long")) {
                    val = mSheet.getCellLongValue(row, mColValIndex);
                } else if (Util.equals(type, "double")) {
                    val = mSheet.getCellDoubleValue(row, mColValIndex);
                } else {
                    if (customTypeReader == null) {
                        val = null;
                    }
                    mSheet.getCell(row, mColValIndex).setCellType(Cell.CELL_TYPE_STRING);
                    String valTmp = mSheet.getCellStringValue(row, mColValIndex);
                    if (type.endsWith("[]")) {
                        val = customTypeReader.readCustomObjArray(type.substring(0, type.length()-2), valTmp);
                    } else {
                        val = customTypeReader.readCustomObj(type, valTmp);
                    }
                }
                caseBean.addDataItem(new ValueBean(type, key, val));
//                System.out.println("type:"+type+"  key:"+key+"  val:"+ val);
            }
        }
        return caseBean;
    }

    /**
     * 返回某一case的JSON数据格式
     * @param index
     * @return
     */
    public JSONObject caseToJson(int index) {
        CaseBean caseBean = readCase(index, null);
        if (caseBean == null) { return null; }
        JSONObject jsonObject = new JSONObject();
        try {
            for (ValueBean data : caseBean.dataList) {
                jsonObject.put(data.key, data.val);
            }
        } catch (JSONException e) {
            e.printStackTrace();
        }
        return jsonObject;
    }

    /**
     * 获取总case数
     * @return
     */
    public int getCaseNum() {
        return mCaseNum;
    }

    /**
     * 获得当前sheet名称
     * @return
     */
    public String getSheetName() {
        if (mSheet == null && mSheet.getSheet() == null) { return null; }
        return mSheet.getSheet().getSheetName();
    }

    public static interface CustomTypeReader {
        public CaseBean readCustomObj(String type, String val);

        public ArrayList<CaseBean> readCustomObjArray(String type, String val);
    }
}
