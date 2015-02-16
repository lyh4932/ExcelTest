package excelOperation;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * @author luoyuheng
 *
 */
public class MySheet{
    private Sheet mSheet;
    private int mergedRegionCnt;
    private ArrayList<CellRangeAddress> mergedRegionArray;

    public MySheet(Sheet sheet) {
        mSheet = sheet;
        if (mSheet != null) {
            mergedRegionCnt = mSheet.getNumMergedRegions();
            mergedRegionArray = new ArrayList<CellRangeAddress>();
            for (int i = 0; i < mergedRegionCnt; i++) {
                mergedRegionArray.add(mSheet.getMergedRegion(i));
            }
        }
    }
    /**
     * judgment if the cell belongs to any merge regions
     * 
     * @param cell
     * @return Return the CellRangeAddress Object the cell belongs to; if
     *         null,this cell isn't belong to any merge regions
     */
    public boolean isInMergeRegion(Cell cell) {
        if (mSheet == null || cell == null) { return false; }
        int rowInd = cell.getRowIndex();
        int colInd = cell.getColumnIndex();
        for (int i = 0; i < mergedRegionCnt; i++) {
            if (mergedRegionArray.get(i).isInRange(rowInd, colInd)) {
                return true;
            }
        }
        return false;
    }

    /**
     * Obtain the CellRangeAddress Object from the merged regions
     * 
     * @param cell
     * @return Return the CellRangeAddress Object the cell belongs to; if
     *         null,this cell isn't belong to any merge regions
     */
    public CellRangeAddress getMergeRegionIndex(Cell cell) {
        if (mSheet == null || cell == null) { return null; }
        int rowId = cell.getRowIndex();
        int colId = cell.getColumnIndex();
        for (int i = 0; i < mergedRegionCnt; i++) {
            CellRangeAddress cellRangeAddress = mergedRegionArray.get(i);
            if (cellRangeAddress.isInRange(rowId, colId)) {
//                System.out.println(cellRangeAddress.formatAsString());
                return cellRangeAddress;
            }
        }
        return null;
    }

    /**
     * return the string value of the cell
     * 
     * @param sheet
     * @param rowId
     * @param colId
     * @return
     */
    public String getCellStringValue(int rowId, int colId) {
        Cell cell = getCell(rowId, colId);
        if (cell == null) { return null; }
        switch (cell.getCellType()) {
        case Cell.CELL_TYPE_BLANK:
            CellRangeAddress cellRangeAddress = getMergeRegionIndex(cell);
            if (cellRangeAddress != null) {
                return getCellStringValue(cellRangeAddress.getFirstRow(),
                        cellRangeAddress.getFirstColumn());
            } else {
                return null;
            }
        case Cell.CELL_TYPE_BOOLEAN:
            return String.valueOf(cell.getBooleanCellValue());
        case Cell.CELL_TYPE_ERROR:
            return null;
        case Cell.CELL_TYPE_FORMULA:
            return String.valueOf(cell.getNumericCellValue());
        case Cell.CELL_TYPE_NUMERIC:
            return String.valueOf(cell.getNumericCellValue());
        case Cell.CELL_TYPE_STRING:
            return cell.getStringCellValue();
        default:
            break;
        }
        return null;
    }

    public boolean getCellBooleanValue(int rowId, int colId) {
        Cell cell = getCell(rowId, colId);
        if (cell == null) { return false; }
        if (cell.getCellType() != Cell.CELL_TYPE_BOOLEAN) {
            return false;
        }
        return cell.getBooleanCellValue();
    }

    public long getCellLongValue(int rowId, int colId) {
        return (long) getCellDoubleValue(rowId, colId);
    }

    public int getCellIntValue(int rowId, int colId) {
        return (int) getCellDoubleValue(rowId, colId);
    }

    public double getCellDoubleValue(int rowId, int colId) {
        Cell cell = getCell(rowId, colId);
        if (cell == null) { return 0; }
        if (cell.getCellType() != Cell.CELL_TYPE_NUMERIC && cell.getCellType() != Cell.CELL_TYPE_FORMULA) {
            return 0;
        }
        return cell.getNumericCellValue();
    }

    /**
     * Obtain the cell object
     * 
     * @param sheet
     * @param rowId
     * @param colId
     * @return
     */
    public Cell getCell(int rowId, int colId) {
        if (mSheet == null || mSheet.getRow(rowId) == null) { return null; }
        return mSheet.getRow(rowId).getCell(colId);
    }

    public void setCellValue(int rowId, int colId, String val) {
        if (mSheet == null) { return; }
        Row row= mSheet.getRow(rowId);
        if (row == null) {
            mSheet.createRow(rowId);
        }
        Cell cell = row.getCell(colId);
        if (cell == null) {
            cell = row.createCell(colId);
        }
        cell.setCellValue(val);
    }

    public Sheet getSheet() {
        return mSheet;
    }
}
