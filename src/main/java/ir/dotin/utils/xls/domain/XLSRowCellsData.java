package ir.dotin.utils.xls.domain;

import java.io.Serializable;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by r.rastakfard on 7/14/2016.
 */
public class XLSRowCellsData implements Serializable {
    Map<String, Integer> columnsDataSize = new HashMap<String, Integer>();
    private Map<String, List<Map<String, String>>> row;
    private int maxHeight;

    public XLSRowCellsData(Map<String, List<Map<String, String>>> row, int maxHeight) {
        this.row = row;
        this.maxHeight = maxHeight;
    }

    public Map<String, List<Map<String, String>>> getRow() {
        return row;
    }

    public void setRow(Map<String, List<Map<String, String>>> row) {
        this.row = row;
    }

    public int getMaxHeight() {
        return maxHeight;
    }

    public void setMaxHeight(int maxHeight) {
        this.maxHeight = maxHeight;
    }

    public Map<String, Integer> getColumnsDataSize() {
        if (columnsDataSize.isEmpty()) {
            for (String key : row.keySet()) {
                columnsDataSize.put(key, row.get(key).size());
            }
        }
        return columnsDataSize;
    }
}
