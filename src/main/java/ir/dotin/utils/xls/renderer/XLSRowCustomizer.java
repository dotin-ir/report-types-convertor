package ir.dotin.utils.xls.renderer;

import ir.dotin.utils.xls.domain.XLSCellStyle;
import ir.dotin.utils.xls.domain.XLSSheetContext;

import java.util.List;
import java.util.Map;

/**
 * Created by r.rastakfard on 7/12/2016.
 */
public interface XLSRowCustomizer {
    Map<String, XLSCellStyle> createRecordStyle(Map<String, List<Map<String, String>>> rowData, XLSSheetContext sheetContext);
}
