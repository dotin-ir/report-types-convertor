package ir.dotin.utils.xls.renderer;

import ir.dotin.utils.xls.domain.*;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by r.rastakfard on 7/17/2016.
 */
public class XLSDefaultRowCustomizer implements XLSRowCustomizer {
    public static XLSCellStyle getDefaultCellStyle(XLSSheetContext sheetContext) {
        boolean even = (sheetContext.getProcessedEntityCount()) % 2 == 0;
        String colorKey = even ? XLSConstants.DEFAULT_EVEN_ROW_COLOR_KEY : XLSConstants.DEFAULT_ODD_ROW_COLOR_KEY;
        XLSColorDescription bgColorDescription = sheetContext.getColorDescription(colorKey);
        XLSColorDescription fontColorDescription = sheetContext.getColorDescription(XLSConstants.DEFAULT_FONT_COLOR_KEY);
        return new XLSCellStyle(new XLSCellFont(fontColorDescription, XLSConstants.DEFAULT_FONT_NAME), bgColorDescription, "TEXT");
    }

    public Map<String, XLSCellStyle> createRecordStyle(Map<String, List<Map<String, String>>> row, XLSSheetContext sheetContext) {
        Map<String, XLSCellStyle> result = new HashMap<String, XLSCellStyle>();
        for (String colName : row.keySet()) {
            XLSCellStyle cellDesign = XLSDefaultRowCustomizer.getDefaultCellStyle(sheetContext);
            result.put(colName, cellDesign);
        }
        return result;
    }
}
