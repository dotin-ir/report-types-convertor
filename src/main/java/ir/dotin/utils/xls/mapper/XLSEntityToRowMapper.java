package ir.dotin.utils.xls.mapper;

import ir.dotin.utils.xls.domain.XLSCellStyle;
import ir.dotin.utils.xls.domain.XLSRowCellsData;
import ir.dotin.utils.xls.domain.XLSSheetContext;

import java.util.Map;

/**
 * Created by r.rastakfard on 7/2/2016.
 */
public interface XLSEntityToRowMapper<E> {
    XLSRowCellsData map(E entity, XLSSheetContext sheetContext);// throws Exception;

    Map<String, XLSCellStyle> getRecordDesign(E entity, XLSSheetContext sheetContext);
}
