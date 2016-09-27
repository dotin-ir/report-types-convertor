package ir.dotin.utils.xls.validator;

import ir.dotin.utils.xls.domain.XLSRecord;
import ir.dotin.utils.xls.domain.XLSSheetContext;

/**
 * Created by r.rastakfard on 6/22/2016.
 */
public class XLSDefaultDocumentValidator extends XLSDocumentValidator {

    @Override
    public boolean validateDocument(XLSSheetContext sheetContext) throws Exception {
        return true;
    }

    @Override
    public boolean validateRow(XLSRecord currentXlsRecord, XLSSheetContext sheetContext) {
        return true;
    }
}
