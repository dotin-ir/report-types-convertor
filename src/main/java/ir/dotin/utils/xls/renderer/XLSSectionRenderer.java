package ir.dotin.utils.xls.renderer;

import ir.dotin.utils.xls.domain.XLSReportSection;
import ir.dotin.utils.xls.domain.XLSSheetContext;

/**
 * Created by r.rastakfard on 7/14/2016.
 */
public interface XLSSectionRenderer<S extends XLSReportSection> {
    void renderSection(XLSSheetContext sheetContext, S section);
}
