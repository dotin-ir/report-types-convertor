package ir.dotin.utils.xls.validator;

import ir.dotin.utils.xls.domain.XLSRecord;
import ir.dotin.utils.xls.domain.XLSSheetContext;

/**
 * Created by r.rastakfard on 6/21/2016.
 */
public abstract class XLSDocumentValidator {
    public abstract boolean validateDocument(XLSSheetContext sheetContext) throws Exception;

    public abstract boolean validateRow(XLSRecord currentRecord, XLSSheetContext sheetContext);

    public boolean defaultValidation(XLSSheetContext xlsSheetContext) {
        return xlsSheetContext.isHeaderOK();
    }
}
