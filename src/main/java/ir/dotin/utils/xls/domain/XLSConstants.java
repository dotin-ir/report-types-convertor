package ir.dotin.utils.xls.domain;

import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * Created by r.rastakfard on 7/17/2016.
 */
public interface XLSConstants {
    String DEFAULT_FONT_NAME = XLSPersianFonts.B_NAZANIN;
    int MAX_COLORS_PER_SHEET = 12;
    String DEFAULT_FONT_COLOR_KEY = "black";
    short DEFAULT_FONT_COLOR_INDEX = IndexedColors.BLACK.getIndex();
    short DEFAULT_ODD_ROW_COLOR_INDEX = 10;
    short DEFAULT_EVEN_ROW_COLOR_INDEX = 11;
    String ERROR_KEY = "errorReason";
    String DEFAULT_ODD_ROW_COLOR_KEY = "LIGHT_GREEN";
    String DEFAULT_EVEN_ROW_COLOR_KEY = "LIGHTER_GREEN";
    String DEFAULT_CELL_FORMAT = "TEXT";
    int DEFAULT_CELL_FONT_SIZE = 10;
    int DEFAULT_ROW_SIZE = 255;

    String DEFAULT_EMPTY_RECORDS_MESSAGE = "DEFAULT_EMPTY_RECORDS_MESSAGE";
    String DEFAULT_EMPTY_SECTION_RECORDS_MESSAGE = "DEFAULT_EMPTY_SECTION_RECORDS_MESSAGE";
    String DEFAULT_ODD_ROW_COLOR_DESCRIPTION = "DEFAULT_ODD_ROW_COLOR_DESCRIPTION";
    String DEFAULT_EVEN_ROW_COLOR_DESCRIPTION = "DEFAULT_EVEN_ROW_COLOR_DESCRIPTION";
    String DEFAULT_FONT_COLOR_DESCRIPTION = "DEFAULT_FONT_COLOR_DESCRIPTION";
    String DEFAULT_COLORS_DESCRIPTION_TABLE_TITLE = "DEFAULT_COLORS_DESCRIPTION_TABLE_TITLE";
    String DEFAULT_REPORT_CONDITIONS_TABLE_TITLE = "DEFAULT_REPORT_CONDITIONS_TABLE_TITLE";
    String INVALID_RECORD = "INVALID_RECORD";
    String EXCEPTION_OCCURRED_IN_MAPPING ="EXCEPTION_OCCURRED_IN_MAPPING";
    String RECORDS_WITH_BAD_FORMAT = "RECORDS_WITH_BAD_FORMAT";
}
