package ir.dotin.utils.xls.checker;

import ir.dotin.utils.xls.domain.XLSDuplicatePolicy;
import ir.dotin.utils.xls.domain.XLSRecord;
import ir.dotin.utils.xls.domain.XLSSheetContext;

/**
 * Created by r.rastakfard on 6/23/2016.
 */
public class XLSDefaultUniqueChecker extends XLSUniqueChecker {

    @Override
    public boolean checkUniqueRow(XLSRecord currentXlsRecord, XLSSheetContext sheetContext) {
        return true;
    }

    @Override
    public XLSDuplicatePolicy getDuplicatePolicy() {
        return XLSDuplicatePolicy.KEEP_ALL;
    }

    public String getHash(XLSRecord currentXlsRecord, XLSSheetContext sheetContext) {
        return "";
    }
}
