package excelOperation;
import java.util.ArrayList;


public class CaseBean {

    public ArrayList<ValueBean> dataList;

    public CaseBean() {
        dataList = new ArrayList<ValueBean>();
    }

    public void addDataItem(ValueBean valueBean){
        dataList.add(valueBean);
    }
}
