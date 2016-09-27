package ir.dotin.utils.xls.writer;

import ir.dotin.utils.xls.domain.*;
import ir.dotin.utils.xls.renderer.XLSDefaultRowCustomizer;
import ir.dotin.utils.xls.renderer.XLSRowCustomizer;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontCharset;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by r.rastakfard on 7/14/2016.
 */
public abstract class XLSBaseWriter {

    protected HashMap<Integer, XLSSheetContext> sheetsContext = new HashMap<Integer, XLSSheetContext>();
    private Map<XLSCellFont, HSSFFont> workbookFonts;
    private XLSRowCustomizer defaultRowCustomizer = new XLSDefaultRowCustomizer();

    public XLSSheetContext getSheetContext(Integer sheetNo) {
        XLSSheetContext context = sheetsContext.get(sheetNo);
        if (context == null) {
            context = new XLSSheetContext(sheetNo);
            sheetsContext.put(sheetNo, context);
        }
        return context;
    }

    protected Collection<XLSSheetContext> getSheetContextsValues() {
        return sheetsContext.values();
    }

    protected void createSheet(HSSFWorkbook workbook, int sheetIndex, XLSSheetContext sheetContext) {
        sheetContext.setWorkbook(workbook);
        sheetContext.initColorsDescriptionList();
        String sheetName = sheetContext.getSheetName();
        HSSFSheet sheet = workbook.createSheet(String.valueOf(StringUtils.isNotEmpty(sheetName) ? sheetName : sheetIndex));
        sheet.setRightToLeft(sheetContext.isRightToLeft());
        sheetContext.setRealSheet(sheet);
    }

    public void setSheetName(int sheetNo, String sheetName) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.setSheetName(sheetName);
    }

    public CellStyle createPOIDataRowStyle(XLSSheetContext sheetContext, boolean isColorful,
                                           boolean hasTopBorder, boolean hasBottomBorder, boolean hasLeftBorder,
                                           boolean hasRightBorder, short bgColor, CellStyle style) {
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(isColorful ? IndexedColors.WHITE.getIndex() : bgColor);
        style.setAlignment(CellStyle.ALIGN_RIGHT);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setDataFormat(HSSFDataFormat.getBuiltinFormat(XLSConstants.DEFAULT_CELL_FORMAT));
        style.setWrapText(true);
        if (hasBottomBorder) {
            style.setBorderBottom(CellStyle.BORDER_THIN);
            style.setBottomBorderColor(IndexedColors.BLUE_GREY.getIndex());
        }
        if (hasLeftBorder) {
            style.setBorderLeft(CellStyle.BORDER_THIN);
            style.setLeftBorderColor(IndexedColors.BLUE_GREY.getIndex());
        }
        if (hasRightBorder) {
            style.setBorderRight(CellStyle.BORDER_THIN);
            style.setRightBorderColor(IndexedColors.BLUE_GREY.getIndex());
        }
        if (hasTopBorder) {
            style.setBorderTop(CellStyle.BORDER_THIN);
            style.setTopBorderColor(IndexedColors.BLUE_GREY.getIndex());
        }
        style.setFont(createPOIStyleFont(sheetContext,XLSDefaultRowCustomizer.getDefaultCellStyle(sheetContext)));
        return style;
    }

    public CellStyle fillCellStyle(XLSSheetContext sheetContext, XLSCellStyle xlsCellStyle, CellStyle style) {
        XLSColorDescription colorDescription = xlsCellStyle.getBackGroundColor();
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(colorDescription.getColorIndex());
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setDataFormat(HSSFDataFormat.getBuiltinFormat(xlsCellStyle.getFormat()));
        style.setWrapText(true);
        if (xlsCellStyle.getHasBottomBorder()) {
            style.setBorderBottom(CellStyle.BORDER_THIN);
            style.setBottomBorderColor(IndexedColors.BLUE_GREY.getIndex());
        }
        if (xlsCellStyle.getHasLeftBorder()) {
            style.setBorderLeft(CellStyle.BORDER_THIN);
            style.setLeftBorderColor(IndexedColors.BLUE_GREY.getIndex());
        }
        if (xlsCellStyle.getHasRightBorder()) {
            style.setBorderRight(CellStyle.BORDER_THIN);
            style.setRightBorderColor(IndexedColors.BLUE_GREY.getIndex());
        }
        if (xlsCellStyle.getHasTopBorder()) {
            style.setBorderTop(CellStyle.BORDER_THIN);
            style.setTopBorderColor(IndexedColors.BLUE_GREY.getIndex());
        }
        style.setFont(createPOIStyleFont(sheetContext,xlsCellStyle));

        return style;
    }

    public HSSFFont createPOIStyleFont(XLSSheetContext sheetContext, XLSCellStyle cellStyle) {
        HSSFFont font;
        Map<XLSCellFont, HSSFFont> workbookFonts = getWorkbookFonts();
        if (!workbookFonts.containsKey(cellStyle.getFont())) {
            XLSCellFont cellFont = cellStyle.getFont();
            HSSFWorkbook workbook = sheetContext.getWorkbook();
            font = workbook.createFont();
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);
            font.setFontHeightInPoints((short) cellFont.getSize());
            font.setFontName(cellFont.getFontName());
            font.setCharSet(FontCharset.ARABIC.getValue());
            font.setColor(cellFont.getFontColor().getColorIndex());
            workbookFonts.put(cellStyle.getFont(),font);
        }else{
            font = workbookFonts.get(cellStyle.getFont());
        }

        return font;
    }

    public CellStyle createPOIHeaderRowStyle(XLSSheetContext sheetContext, short bgColor, boolean hasTopBorder,
                                             boolean hasBottomBorder, boolean hasLeftBorder, boolean hasRightBorder, CellStyle style) {
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(bgColor);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setDataFormat(HSSFDataFormat.getBuiltinFormat(XLSConstants.DEFAULT_CELL_FORMAT));
        if (hasBottomBorder) {
            style.setBorderBottom(CellStyle.BORDER_THIN);
            style.setBottomBorderColor(IndexedColors.BLUE_GREY.getIndex());
        }
        if (hasLeftBorder) {
            style.setBorderLeft(CellStyle.BORDER_THIN);
            style.setLeftBorderColor(IndexedColors.BLUE_GREY.getIndex());
        }
        if (hasRightBorder) {
            style.setBorderRight(CellStyle.BORDER_THIN);
            style.setRightBorderColor(IndexedColors.BLUE_GREY.getIndex());
        }
        if (hasTopBorder) {
            style.setBorderTop(CellStyle.BORDER_THIN);
            style.setTopBorderColor(IndexedColors.BLUE_GREY.getIndex());
        }
        HSSFFont font = createPOIStyleFont(sheetContext, XLSDefaultRowCustomizer.getDefaultCellStyle(sheetContext));
        style.setFont(font);
        return style;
    }

    public CellStyle createPOIHeaderRowStyle(XLSSheetContext sheetContext, short bgColor, boolean hasTopBorder,
                                             boolean hasBottomBorder, boolean hasLeftBorder, boolean hasRightBorder, int fontSize) {
        HSSFCellStyle style = sheetContext.getWorkbook().createCellStyle();
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(bgColor);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setDataFormat(HSSFDataFormat.getBuiltinFormat(XLSConstants.DEFAULT_CELL_FORMAT));
        if (hasBottomBorder) {
            style.setBorderBottom(CellStyle.BORDER_THIN);
            style.setBottomBorderColor(IndexedColors.BLUE_GREY.getIndex());
        }
        if (hasLeftBorder) {
            style.setBorderLeft(CellStyle.BORDER_THIN);
            style.setLeftBorderColor(IndexedColors.BLUE_GREY.getIndex());
        }
        if (hasRightBorder) {
            style.setBorderRight(CellStyle.BORDER_THIN);
            style.setRightBorderColor(IndexedColors.BLUE_GREY.getIndex());
        }
        if (hasTopBorder) {
            style.setBorderTop(CellStyle.BORDER_THIN);
            style.setTopBorderColor(IndexedColors.BLUE_GREY.getIndex());
        }
        XLSCellStyle defaultCellStyle = XLSDefaultRowCustomizer.getDefaultCellStyle(sheetContext);
        defaultCellStyle.getFont().setSize(fontSize);
        HSSFFont font = createPOIStyleFont(sheetContext, defaultCellStyle);
        style.setFont(font);
        return style;
    }

    public FileOutputStream createDocument(String filePath) throws IOException {
        ByteArrayOutputStream outputStream = createDocument();
        FileOutputStream fileOutputStream = new FileOutputStream(filePath);
        fileOutputStream.write(outputStream.toByteArray());
        fileOutputStream.flush();
        fileOutputStream.close();
        return fileOutputStream;
    }

    public XLSRowCustomizer getDefaultRowCustomizer() {
        return defaultRowCustomizer;
    }

    public abstract ByteArrayOutputStream createDocument() throws IOException;

    public void setRawPOIRecords(int sheetNo, Map<Integer, List<HSSFRow>> badRowsMap) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.setRawPOIRecords(badRowsMap);
    }

    public void addBusinessVariable(int sheetNo, String variableName, Object variableValue) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.addBusinessVariable(variableName, variableValue);
    }

    public Map<XLSCellFont, HSSFFont> getWorkbookFonts() {
        if (workbookFonts==null){
            workbookFonts = new HashMap<XLSCellFont, HSSFFont>();
        }
        return workbookFonts;
    }

    public abstract void createReportColorsDescriptionTable(XLSSheetContext sheetContext);

    public int computeRowHeight(String value, Integer width) {
        int length = value.length();
        int totalCharInMergedCells = width * 5;
        int rowCount = 1;
        if (length > totalCharInMergedCells) {
            rowCount++;
            int remaining = length - totalCharInMergedCells;
            while (remaining >= totalCharInMergedCells) {
                remaining -= totalCharInMergedCells;
                rowCount++;
            }
        }
        return rowCount * XLSConstants.DEFAULT_ROW_SIZE;
    }

    public void addColorsDescription(int sheetNo, String colorKey, int red, int green, int blue, String description) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.addColorToSheet(new XLSColorDescription(colorKey, red, green, blue, description));
    }
}
