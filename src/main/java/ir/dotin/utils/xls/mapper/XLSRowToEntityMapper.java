package ir.dotin.utils.xls.mapper;

import ir.dotin.utils.xls.domain.XLSRecord;
import ir.dotin.utils.xls.domain.XLSSheetContext;

/**
 * Created by r.rastakfard on 6/29/2016.
 */
public interface XLSRowToEntityMapper<E> {
    E map(XLSRecord record, XLSSheetContext sheetContext) throws Exception;
}
