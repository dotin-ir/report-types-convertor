package ir.dotin.utils.xls.checker;

import ir.dotin.utils.xls.domain.XLSDuplicatePolicy;
import ir.dotin.utils.xls.domain.XLSRecord;
import ir.dotin.utils.xls.domain.XLSSheetContext;

import java.util.List;
import java.util.Map;

/**
 * Created by r.rastakfard on 6/30/2016.
 */
public class XLSPrimaryKeyChecker extends XLSUniqueChecker {

    private String primaryKey;

    public XLSPrimaryKeyChecker(String primaryKey) {
        this.primaryKey = primaryKey;
    }

    public boolean checkUniqueRow(XLSRecord currentXlsRecord, XLSSheetContext sheetContext) {
        onFinish(true, sheetContext);
        return true;
    }

    public XLSDuplicatePolicy getDuplicatePolicy() {
        return XLSDuplicatePolicy.REPORT_ALL_AS_ERROR;
    }

    public String getHash(XLSRecord currentXlsRecord, XLSSheetContext sheetContext) {
        List<Map<String, String>> maps = currentXlsRecord.getRecordData().get(primaryKey);
        if (maps == null) {
            throw new RuntimeException("Data does not exist in record for primary key : " + primaryKey);
        }
        if (maps.size() > 1) {
            throw new RuntimeException("Can not compute hash for list column(you can not set list column as unique key - " + primaryKey + ")!");
        }
        String data = maps.get(0).get(primaryKey);
        if (data != null) {
            return String.valueOf(data.hashCode());
        } else {
            return primaryKey;
        }

    }
}
