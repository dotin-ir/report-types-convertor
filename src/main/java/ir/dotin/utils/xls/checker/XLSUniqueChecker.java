package ir.dotin.utils.xls.checker;

import ir.dotin.utils.xls.domain.XLSColumnDefinition;
import ir.dotin.utils.xls.domain.XLSDuplicatePolicy;
import ir.dotin.utils.xls.domain.XLSRecord;
import ir.dotin.utils.xls.domain.XLSSheetContext;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Created by r.rastakfard on 6/23/2016.
 */
public abstract class XLSUniqueChecker {

    public int EQUAL = 0;
    public int GREATER = 1;
    public int LOWER = -1;
    private Boolean keySet = null;
    private boolean status = true;

    public abstract boolean checkUniqueRow(XLSRecord currentXlsRecord, XLSSheetContext sheetContext);

    public abstract XLSDuplicatePolicy getDuplicatePolicy();

    public abstract String getHash(XLSRecord currentXlsRecord, XLSSheetContext sheetContext);

    public int compareRecord(XLSRecord record1, XLSRecord record2, XLSSheetContext context) {
        Map<String, List<Map<String, String>>> recordData1 = record1.getRecordData();
        Map<String, List<Map<String, String>>> recordData2 = record2.getRecordData();
        if (isKeySet(context)) {
            List<XLSColumnDefinition> recordColumns1 = record1.getRecordColumns();
            List<XLSColumnDefinition> recordColumns2 = record2.getRecordColumns();
            checkColumnMatch(recordColumns1, recordColumns2);
            for (XLSColumnDefinition definition : recordColumns1) {
                if (definition.isUniqueColumn() && definition.isRealColumn()) {
                    List<Map<String, String>> colDataList = recordData1.get(definition.getName());
                    List<Map<String, String>> colDataList1 = recordData2.get(definition.getName());
                    if (colDataList == null) colDataList = new ArrayList<Map<String, String>>();
                    if (colDataList1 == null) colDataList1 = new ArrayList<Map<String, String>>();
                    if (colDataList.size() > 1 || colDataList1.size() > 1) {
                        throw new RuntimeException("you can not use column with data list as unique column!");
                    }
                    return XLSUtils.compareSimpleRecord(colDataList.get(0), colDataList1.get(0));
                }
            }
               /* String recordDataForKey1 = recordData1.get(uniqueKey);
                String recordDataForKey2 = recordData2.get(uniqueKey);
                if (recordDataForKey1 == null && recordDataForKey2 != null) {
                    return LOWER;
                }
                if (recordDataForKey1 != null && recordDataForKey2 == null) {
                    return GREATER;
                }
                if (recordDataForKey1 != null && recordDataForKey2 != null) {
                    //TODO
                    int lastIndexOf1 = recordDataForKey1.lastIndexOf(".");
                    if (lastIndexOf1 != -1) {
                        recordDataForKey1 = recordDataForKey1.substring(0, lastIndexOf1);
                    }

                    int lastIndexOf2 = recordDataForKey2.lastIndexOf(".");
                    if (lastIndexOf2 != -1) {
                        recordDataForKey2 = recordDataForKey2.substring(0, lastIndexOf2);
                    }
                    int compareResult = recordDataForKey1.compareTo(recordDataForKey2);
                    if (compareResult != EQUAL) {
                        return compareResult;
                    }
                }*/
            return EQUAL;
        }
        return record1.compareTo(record2);
    }

    private void checkColumnMatch(List<XLSColumnDefinition> def1, List<XLSColumnDefinition> def2) {

    }

    public boolean isKeySet(XLSSheetContext context) {
        return context.isUniqueColumnKeySet();
    }

    public boolean getStatus() {
        return status;
    }

    public void setStatus(boolean status) {
        this.status = status;
    }

    public void onFinish(boolean valid, XLSSheetContext sheetContext) {
        setStatus(valid);
        sheetContext.setUniqueCheckerStatus(valid);
    }
}
