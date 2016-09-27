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

    String DEFAULT_EMPTY_RECORDS_MESSAGE = "داده ای وجود ندارد";
    String DEFAULT_EMPTY_SECTION_RECORDS_MESSAGE = "با شرایط فوق، داده ای وجود ندارد";
    String DEFAULT_ODD_ROW_COLOR_DESCRIPTION = "رنگ پیش فرض ردیف های فرد";
    String DEFAULT_EVEN_ROW_COLOR_DESCRIPTION = "رنگ پیش فرض ردیف های زوج";
    String DEFAULT_FONT_COLOR_DESCRIPTION = "رنگ پیش فرض فونت";
    String DEFAULT_COLORS_DESCRPTION_TABLE_TITLE = "توضیحات رنگ ها";
    String EXCEPTION_OCCURED_IN_MAPPING="در هنگام نگاشت موجودیت خطایی رخ داده است";
}
