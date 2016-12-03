package ir.dotin.utils.xls.writer;

import ir.dotin.utils.xls.checker.XLSUtils;
import ir.dotin.utils.xls.domain.*;
import ir.dotin.utils.xls.renderer.XLSDefaultRowCustomizer;
import ir.dotin.utils.xls.renderer.XLSRowCustomizer;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.*;

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
public abstract class XLSBaseWriter<B> {

    protected HashMap<Integer, XLSSheetContext> sheetsContext = new HashMap<Integer, XLSSheetContext>();
    private Map<XLSCellFont, HSSFFont> workbookFonts;
    private XLSRowCustomizer defaultRowCustomizer = new XLSDefaultRowCustomizer();
    private HSSFWorkbook workBook;
    private String basicInfoSheetProtectionPassword;


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

    protected void createSheet(int sheetIndex, XLSSheetContext sheetContext) {
        sheetContext.setWorkbook(getWorkBook());
        sheetContext.initColorsDescriptionList();
        String sheetName = sheetContext.getSheetName();
        HSSFSheet sheet = getWorkBook().createSheet(String.valueOf(StringUtils.isNotEmpty(sheetName) ? sheetName : sheetIndex));
        sheet.setRightToLeft(sheetContext.isRightToLeft());
        if (StringUtils.isNotEmpty(sheetContext.getReadOnlyPassword())) {
            sheet.protectSheet(sheetContext.getReadOnlyPassword());
        }
        sheetContext.setSheetIndex(sheetIndex);
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
        style.setFont(createPOIStyleFont(sheetContext, XLSDefaultRowCustomizer.getDefaultCellStyle(sheetContext)));
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
        style.setFont(createPOIStyleFont(sheetContext, xlsCellStyle));

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
            workbookFonts.put(cellStyle.getFont(), font);
        } else {
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
        if (workbookFonts == null) {
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

    public void setSheetAsReadonly(int sheetIndex, String password) {
        XLSSheetContext sheetContext = getSheetContext(sheetIndex);
        if (sheetContext != null) {
            if (StringUtils.isEmpty(password)) {
                throw new IllegalArgumentException(XLSUtils.getProperty(XLSConstants.INVALID_EMPTY_PASSWORD));
            }
            sheetContext.setReadOnlyPassword(password);
        }
    }

    private Map<String, XLSBasicInfo> basiceInfos;

    public Map<String, XLSBasicInfo> getBasiceInfos() {
        if (basiceInfos == null) {
            basiceInfos = new HashMap<String, XLSBasicInfo>();
        }
        return basiceInfos;
    }

    public void setBasiceInfos(Map<String, XLSBasicInfo> basiceInfos) {
        this.basiceInfos = basiceInfos;
    }

    public void addBasicInfoCollection(String basicInfoKey, String errorBoxTitle, String errorBoxMessage, List<B> collection) {
        if (StringUtils.isNotEmpty(basicInfoKey)) {
            if (getBasiceInfos().containsKey(basicInfoKey)) {
                throw new IllegalArgumentException("Basic ifo with key '" + basicInfoKey + "' already exist!");
            }
            if (collection != null) {
                XLSBasicInfo basicInfo = new XLSBasicInfo(basicInfoKey, errorBoxTitle, errorBoxMessage, collection);
                getBasiceInfos().put(basicInfoKey, basicInfo);
            }
        } else {
            throw new IllegalArgumentException("Basic ifo key is empty!");
        }
    }

    protected int generateSheetIndex(XLSSheetContext sheetContext, int sheetIndex) {
        return sheetIndex + 1;
    }

    protected void refreshBasicInfoBasedOnSheetContexts() {
        for (XLSSheetContext sheetContext : getSheetContextsValues()) {
            List<XLSColumnDefinition> columnsDefinition = getSheetContextColumnsDefinition(sheetContext);
            int basicInfoCount = 0;
            for (XLSColumnDefinition definition : columnsDefinition) {
                if (StringUtils.isNotEmpty(definition.getBasicInfoCollectionKey())){
                    if (!getBasiceInfos().containsKey(definition.getBasicInfoCollectionKey()) && definition.getBasicInfoCollection()==null){
                        throw new RuntimeException("Basic information for column '" + definition.getName() + "' has no reference to data!");
                    }
                }
                if (definition.getBasicInfoCollection() != null) {
                    basicInfoCount += 1;
                    String basicInfoCollectionKey = definition.getBasicInfoCollectionKey();
                    if (StringUtils.isEmpty(basicInfoCollectionKey)) {
                        basicInfoCollectionKey = String.valueOf(basicInfoCount);
                    }
                    addBasicInfoCollection(basicInfoCollectionKey, XLSUtils.getProperty(XLSConstants.DEFAULT_BASIC_INFO_SHEET_ERROR_TITLE),
                            XLSUtils.getProperty(XLSConstants.DEFAULT_BASIC_INFO_SHEET_ERROR_MESSAGE), definition.getBasicInfoCollection());
                    List<XLSColumnDefinition> subColumns = definition.getSubColumns();
                    for (XLSColumnDefinition subDefinition : subColumns) {
                        if (subDefinition.getBasicInfoCollection() != null) {
                            basicInfoCount += 1;
                            basicInfoCollectionKey = subDefinition.getBasicInfoCollectionKey();
                            if (StringUtils.isEmpty(basicInfoCollectionKey)) {
                                basicInfoCollectionKey = String.valueOf(basicInfoCount);
                            }
                            addBasicInfoCollection(basicInfoCollectionKey, XLSUtils.getProperty(XLSConstants.DEFAULT_BASIC_INFO_SHEET_ERROR_TITLE),
                                    XLSUtils.getProperty(XLSConstants.DEFAULT_BASIC_INFO_SHEET_ERROR_MESSAGE), subDefinition.getBasicInfoCollection());
                        }
                    }
                }
            }
        }
        Map<String, XLSBasicInfo> basicInfos = getBasiceInfos();
        if (!basicInfos.isEmpty()) {
//            HSSFSheet sheet = getWorkBook().createSheet(XLSConstants.BASIC_INFO_SHEET_NAME);
//            sheet.setRightToLeft(true);
//            sheet.protectSheet(getBasicInfoSheetProtectionPassword());
            int basicInfoCellIndex = 0;
            for (String columnKey : basicInfos.keySet()) {
                XLSBasicInfo basicInfo = basicInfos.get(columnKey);
                basicInfo.setCellIndex(basicInfoCellIndex);
//                Name namedCell = getWorkBook().createName();
//                namedCell.setNameName(columnKey);
//                namedCell.setSheetIndex(0);
               /* namedCell.setRefersToFormula(XLSConstants.BASIC_INFO_SHEET_NAME + "!$" + CellReference.convertNumToColString(basicInfoCellIndex)
                        + "$2:$" + CellReference.convertNumToColString(basicInfoCellIndex) + "$" + (basicInfo.getCollection().size() + 1));
                DVConstraint constraint = DVConstraint.createFormulaListConstraint(XLSConstants.BASIC_INFO_SHEET_NAME + "!$" + CellReference.convertNumToColString(basicInfoCellIndex)
                        + "$2:$" + CellReference.convertNumToColString(basicInfoCellIndex) + "$" + (basicInfo.getCollection().size() + 1));*/
                /*DVConstraint constraint = DVConstraint.createFormulaListConstraint(XLSConstants.BASIC_INFO_SHEET_NAME + "!$" + CellReference.convertNumToColString(basicInfoCellIndex)
                        + "$2:$" + CellReference.convertNumToColString(basicInfoCellIndex) + "$" + (basicInfo.getCollection().size()+1));*/
//                basicInfo.setConstraint(constraint);
//                createBasicInfoColumn(sheet, basicInfoCellIndex, columnKey, basicInfo);
                basicInfoCellIndex++;
            }
//            int totalSheetCount = getTotalSheetCount();
//            getWorkBook().setSheetHidden(totalSheetCount, true);
        }
    }

    protected abstract List<XLSColumnDefinition> getSheetContextColumnsDefinition(XLSSheetContext sheetContext);

    public void createBasicInfoColumn(HSSFSheet sheet, int basicInfoCellIndex, String columnKey, XLSBasicInfo basicInfo) {
        HSSFRow headerRow = getRow(sheet, 0);
        HSSFCell headerCell = headerRow.createCell(basicInfoCellIndex);
        headerCell.setCellValue(columnKey);
        int dataRowIndex = 1;
        List<B> collection = basicInfo.getCollection();
        for (B data : collection) {
            HSSFRow dataRow = getRow(sheet, dataRowIndex);
            dataRow.createCell(basicInfoCellIndex).setCellValue(data.toString());
            dataRowIndex++;
        }
    }

    protected HSSFRow getRow(HSSFSheet sheet, int rowIndex) {
        HSSFRow row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        return row;
    }


    public int getTotalSheetCount() {
        return sheetsContext.size();
    }

    public HSSFWorkbook getWorkBook() {
        if (workBook == null) {
            workBook = new HSSFWorkbook();
        }
        return workBook;
    }

    public void setWorkBook(HSSFWorkbook workBook) {
        this.workBook = workBook;
    }

    public void setBasicInfoSheetProtectionPassword(String basicInfoSheetProtectionPassword) {
        this.basicInfoSheetProtectionPassword = basicInfoSheetProtectionPassword;
    }

    public String getBasicInfoSheetProtectionPassword() {
        if (StringUtils.isEmpty(basicInfoSheetProtectionPassword)) {
            return XLSConstants.DEFAULT_BASIC_INFO_SHEET_PASSWORD;
        }
        return basicInfoSheetProtectionPassword;
    }

    public boolean hasSheet(String sheetName){
        return getWorkBook().getSheet(sheetName)!=null;
    }
}
