package ir.dotin.utils.xls.domain;

import ir.dotin.utils.xls.checker.XLSUniqueChecker;
import ir.dotin.utils.xls.checker.XLSUtils;
import ir.dotin.utils.xls.mapper.XLSEntityToRowMapper;
import ir.dotin.utils.xls.mapper.XLSRowToEntityMapper;
import ir.dotin.utils.xls.renderer.XLSRowCustomizer;
import ir.dotin.utils.xls.validator.XLSDefaultDocumentValidator;
import ir.dotin.utils.xls.validator.XLSDocumentValidator;
import ir.dotin.utils.xls.writer.ExcelReportGenerator;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.Serializable;
import java.util.*;

/**
 * Created by r.rastakfard on 6/22/2016.
 */
public class XLSSheetContext<E> implements Serializable {
    private static String NO_TITLE = "گزارش عنوان ندارد";

    private final int sheetNumber;
    Integer lastSheetRecordIndex = 0;
    private XLSDocumentValidator documentValidator = new XLSDefaultDocumentValidator();
    private List<XLSUniqueChecker> uniqueCheckers;
    private HSSFWorkbook workbook;
    private boolean parsed = false;
    private List<Map<String, String>> invalidRecords = new ArrayList<Map<String, String>>();
    private List<XLSRecord> rawRecords = new ArrayList<XLSRecord>();
    private Map<String, XLSRecord> rawRecordsWithPK = new HashMap<String, XLSRecord>();
    private List objectRecords = new ArrayList();
    private Map<String, Object> objectRecordsWithPK = new HashMap<String, Object>();
    private List<XLSRecord> errorRecords = new ArrayList<XLSRecord>();
    private Set<String> errorRecordsWithKey = new HashSet<String>();
    private XLSRowToEntityMapper resultMapper;
    private boolean uniqueCheckerStatus = true;
    private String primaryKey;
    private List<E> entityRecords = new ArrayList<E>();
    private XLSEntityToRowMapper entityToRowMapper;
    private List<XLSColumnDefinition> columnsDefinition;
    private CellStyle headerColumnsStyle;
    private Map<String, String> parsedHeaders = new HashMap<String, String>();
    private boolean headerParsed = false;
    private Map<Integer, CellStyle> colsStyle = new HashMap<Integer, CellStyle>();
    private String sheetName;
    private boolean rightToLeft = true;
    private XLSRowCustomizer rowCustomizer;
    private String reportTitle;
    private List<XLSReportSection> reportSections;
    private HSSFSheet realSheet;
    private ExcelReportGenerator reportGenerator;
    private List<XLSColorDescription> colorsDescription;
    private Map<String, XLSColorDescription> colorsDescriptionMap;
    private Map<XLSCellStyle, XLSCellStyle> cellStyleMap;
    private int dummyRowsCount;
    private E dummyEntity;
    private Map<Integer, List<HSSFRow>> badFormatRecords;
    private boolean uniqueColumnKeySet = false;
    private boolean uniqueColumnKeyParesed = false;
    private List<String> uniqueKeys;
    private List<XLSReportField> reportInformationFields;
    private Map<String, Object> businessVariables;
    private Map<Integer, List<HSSFRow>> rawPOIRecords;
    private String emptyRecordsMessage = XLSConstants.DEFAULT_EMPTY_RECORDS_MESSAGE;
    private Integer processedEntityCount = 0;
    private String readOnlyPassword;


    public XLSSheetContext(int sheetNumber) {
        this.sheetNumber = sheetNumber;
    }

    public XLSSheetContext(HSSFWorkbook workbook, int sheetNumber) {
        if (workbook == null) {
            throw new RuntimeException("workbook is null, invalid document!");
        }
        this.workbook = workbook;
        this.sheetNumber = sheetNumber;
    }

    public XLSDocumentValidator getDocumentValidator() {
        if (documentValidator == null) {
            documentValidator = new XLSDefaultDocumentValidator();
        }
        return documentValidator;
    }

    public void setDocumentValidator(XLSDocumentValidator documentValidator) {
        this.documentValidator = documentValidator;
    }

    public List<XLSColumnDefinition> getHeaderColumns() {
        return columnsDefinition;
    }

    public void setHeaderColumns(List<XLSColumnDefinition> headerColumns) {
        this.columnsDefinition = headerColumns;
    }

    public List<XLSColumnDefinition> getRealHeaderColumns() {
        List<XLSColumnDefinition> result = new ArrayList<XLSColumnDefinition>();
        for (XLSColumnDefinition definition : columnsDefinition) {
            if (!definition.isRealColumn()) {
                for (XLSColumnDefinition subDef : definition.getSubColumns()) {
                    result.add(subDef);
                }
            } else {
                result.add(definition);
            }
        }
        return result;
    }

    public Map<String, String> getParsedHeaders() {
        return parsedHeaders;
    }

    public List<Map<String, String>> getInvalidRecords() {
        return invalidRecords;
    }

    public void setInvalidRecords(List<Map<String, String>> invalidRecords) {
        this.invalidRecords = invalidRecords;
    }

    public List<XLSUniqueChecker> getUniqueCheckers() {
        if (uniqueCheckers == null) {
            uniqueCheckers = new ArrayList<XLSUniqueChecker>();
        }
        return uniqueCheckers;
    }

    public boolean hasColWithHeaderName(String colName) {
        if (isHeaderOK()) {
            for (String key : getParsedHeaders().keySet()) {
                String headerColValue = getParsedHeaders().get(key);
                if (headerColValue.equals(colName)) {
                    return true;
                }
            }
        } else {
            throw new RuntimeException("Header information does not match header metadata information!");
        }
        /*HSSFRow firstSheetRow = getWorkbook().getSheetAt(getSheetNumber()).getRow(0);
        short lastCellNum = firstSheetRow.getLastCellNum();
        for (int cellIndex = 0; cellIndex < lastCellNum; cellIndex++) {
            HSSFCell cell = firstSheetRow.getCell(cellIndex);
            if (colName.equals(cell.getStringCellValue())) {
                return true;
            }
        }*/
        return false;
    }

    public HSSFWorkbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(HSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    public int getSheetNumber() {
        return sheetNumber;
    }

    public Object getRecords() {
        if (hasPrimaryKey()) {
            if (getResultMapper() != null) {
                return objectRecordsWithPK;
            } else {
                return rawRecordsWithPK;
            }
        } else {
            if (getResultMapper() != null) {
                return objectRecords;
            } else {
                return rawRecords;
            }
        }
    }

    public List<XLSRecord> getRawRecords() {
        if (rawRecords == null) {
            rawRecords = new ArrayList<XLSRecord>();
        }
        return rawRecords;
    }

    public void setRawRecords(List<Map<String, String>> rawRecords) {
        for (Map<String, String> record : rawRecords) {
            addRawRecord(record);
        }
    }

    public void setRawXLSRecords(List<XLSRecord> rawXLSRecords) {
        if (rawXLSRecords != null && !rawXLSRecords.isEmpty()) {
            for (XLSRecord record : rawXLSRecords) {
                getRawRecords().add(record);
            }
        }
    }

    public void addRawRecord(Map<String, String> record) {
        if (record == null || record.isEmpty()) {
            return;
        }
        synchronized (rawRecords) {
            getRawRecords().add(new XLSRecord(record));
        }
    }

    public void addRawEntityRecord(Map<String, List<Map<String, String>>> record) {
        XLSRecord xlsRecord = new XLSRecord();
        xlsRecord.setRecordData(record);
        this.rawRecords.add(xlsRecord);
    }

    public void setRawEntityRecord(List<Map<String, List<Map<String, String>>>> records) {
        for (Map<String, List<Map<String, String>>> record : records) {
            addRawEntityRecord(record);
        }
    }

    public List<XLSRecord> getErrorRecords() {
        return errorRecords;
    }

    public Set<String> getErrorRecordsWithKey() {
        return errorRecordsWithKey;
    }

    public HSSFSheet getSheetInstance() {
        return getWorkbook().getSheetAt(sheetNumber);
    }

    public int getFirstRowIndex() {
        List<XLSColumnDefinition> headerColumns = getHeaderColumns();
        for (XLSColumnDefinition definition : headerColumns) {
            if (hasSubColumns(definition)) {
                return 2;
            }
        }
        return 1;
    }

    public boolean hasSubColumns(XLSColumnDefinition xlsColumnDefinition) {
        return XLSUtils.hasSubColumns(xlsColumnDefinition);
    }

    public void validateDocumentSheet() throws Exception {
        XLSDocumentValidator validator = getDocumentValidator();
        boolean isValidDocument = validator.validateDocument(this);
        if (!isValidDocument) {
            throw new Exception("Invalid Document Content!");
        }
    }


    public void createResultRecordsObject() {
        if (hasPrimaryKey()) {
            if (getResultMapper() != null) {
                List<XLSRecord> tmpRawRecords = new ArrayList<XLSRecord>();
                for (XLSRecord record : tmpRawRecords) {
                    XLSRowToEntityMapper resultMapper = getResultMapper();
                    try {
                        Map<String, List<Map<String, String>>> recordData = record.getRecordData();
                        Object resultRecord = resultMapper.map(record, this);
                        objectRecordsWithPK.put(getRecordPrimaryKeyValue(recordData, getPrimaryKey()), resultRecord);
                    } catch (Exception ex) {
                        record.setSimpleDataValue(XLSConstants.ERROR_KEY, ex.getMessage());
                        tmpRawRecords.add(record);
                    }
                }
                if (!tmpRawRecords.isEmpty()) {
                    errorRecords.addAll(tmpRawRecords);
                    this.rawRecords.removeAll(tmpRawRecords);
                }
            } else {
                for (XLSRecord record : getRawRecords()) {
                    Map<String, List<Map<String, String>>> recordData = record.getRecordData();
                    rawRecordsWithPK.put(getRecordPrimaryKeyValue(recordData, getPrimaryKey()), record);
                }
            }
        } else {
            if (getResultMapper() != null) {
                List<XLSRecord> tmpRawRecords = new ArrayList<XLSRecord>();
                for (XLSRecord record : getRawRecords()) {
                    XLSRowToEntityMapper resultMapper = getResultMapper();
                    try {
                        Object resultRecord = resultMapper.map(record, this);
                        objectRecords.add(resultRecord);
                    } catch (Exception ex) {
                        if (StringUtils.isEmpty(ex.getMessage())) {
                            record.setSimpleDataValue(XLSConstants.ERROR_KEY, XLSConstants.EXCEPTION_OCCURRED_IN_MAPPING);
                        } else {
                            record.setSimpleDataValue(XLSConstants.ERROR_KEY, ex.getMessage());
                        }
                        tmpRawRecords.add(record);
                    }
                }
                if (!tmpRawRecords.isEmpty()) {
                    errorRecords.addAll(tmpRawRecords);
                    this.rawRecords.removeAll(tmpRawRecords);
                }
            }
        }
    }

    private String getRecordPrimaryKeyValue(Map<String, List<Map<String, String>>> recordData, String primaryKey) {
        return "";
    }

    private boolean hasPrimaryKey() {
        return StringUtils.isNotEmpty(getPrimaryKey());
    }

    public void addToErrorList(String key) {
        if (!errorRecordsWithKey.contains(key)) {
            errorRecordsWithKey.add(key);
        }
    }

    public void addToErrorList(String key, XLSRecord currentXlsRecord) {
        if (!errorRecordsWithKey.contains(key)) {
            errorRecordsWithKey.add(key);
        }
    }

    public boolean isExistInErrorRecord(String key) {
        return errorRecordsWithKey.contains(key);
    }

    public String getCellValue(HSSFCell cell) {
        if (cell == null) return "";
        String result = "";
        try {
            result = cell.getStringCellValue();
        } catch (Exception ex) {
            result = String.format("%.0f", cell.getNumericCellValue());
//            result = String.valueOf(cell.getNumericCellValue());
        }
        return result;
        /*int cellType = cell.getCellType();
        if (cellType == 1) {
            return cell.getStringCellValue().trim();
        } else {
            return String.valueOf(cell.getNumericCellValue());
        }*/
    }

    public HSSFRow getRow(int rowNo) {
        return getSheetInstance().getRow(rowNo);
    }

    public boolean isHeaderSet() {
        return columnsDefinition != null && !columnsDefinition.isEmpty();
    }

    private void checkSheetParsed() {
        if (!parsed) {
            throw new RuntimeException("please Parse Sheet first!");
        }
    }

    public HSSFRow getFirstRow() {
        return getSheetInstance().getRow(0);
    }

    public boolean isHeaderOK() {
        if (isHeaderSet() && !headerParsed) {
            List<XLSColumnDefinition> headerColumns = getHeaderColumns();
            Map<String, Integer> columnWithSameHeaderCount = new HashMap<String, Integer>();
            boolean hasSubCols = false;
            HSSFRow firstRow = getRow(0);
            int realHeaderColIndex = 0;
            for (XLSColumnDefinition headerCol : headerColumns) {
                Integer width = headerCol.getRealColumnWidth();
                String headerColValue = getCellValue(firstRow.getCell(realHeaderColIndex));
                if (headerColValue == null) return false;
                if (duplicateColumnFound(columnWithSameHeaderCount, headerCol, headerColValue)) return false;
                parsedHeaders.put(headerCol.getName(), headerCol.getfName());
                if (hasSubColumns(headerCol)) {
                    hasSubCols = true;
                }
                realHeaderColIndex += width;

            }
            if (hasSubCols) {
                HSSFRow subColsRow = getRow(1);
                realHeaderColIndex = 0;
                for (XLSColumnDefinition headerCol : headerColumns) {
                    if (hasSubColumns(headerCol)) {
                        for (XLSColumnDefinition subColDef : headerCol.getSubColumns()) {
                            Integer width = subColDef.getWidth();
                            String subColValue = getCellValue(subColsRow.getCell(realHeaderColIndex));
                            if (subColValue == null) return false;
                            if (duplicateColumnFound(columnWithSameHeaderCount, subColDef, subColValue)) return false;
                            parsedHeaders.put(subColDef.getName(), subColDef.getfName());
                            realHeaderColIndex += width;
                        }
                    } else {
                        Integer width = headerCol.getWidth();
                        realHeaderColIndex += width;
                    }

                }
            }
            headerParsed = true;
        }
        return true;
    }

    private boolean duplicateColumnFound(Map<String, Integer> columnWithSameHeaderCount, XLSColumnDefinition subColDef, String subColValue) {
        if (!subColValue.equals(subColDef.getfName())) {
            return true;
        } else {
            Integer integer = columnWithSameHeaderCount.get(subColDef.getName());
            if (integer == null) {
                columnWithSameHeaderCount.put(subColDef.getName(), 1);
            } else {
                return true;
            }
        }
        return false;
    }

    public XLSRowToEntityMapper getResultMapper() {
        return resultMapper;
    }

    public void setResultMapper(XLSRowToEntityMapper resultMapper) {
        this.resultMapper = resultMapper;
    }

    public List getObjectRecords() {
        return objectRecords;
    }

    public void addUniqueChecker(XLSUniqueChecker uniqueChecker) {
        getUniqueCheckers().add(uniqueChecker);

    }

    public boolean isUniqueCheckerStatus() {
        return uniqueCheckerStatus;
    }

    public void setUniqueCheckerStatus(boolean uniqueCheckerStatus) {
        this.uniqueCheckerStatus = uniqueCheckerStatus;
    }

    public String getPrimaryKey() {
        return primaryKey;
    }

    public void setPrimaryKey(String primaryKey) {
        if (StringUtils.isEmpty(primaryKey)) {
            throw new RuntimeException("primary key cannot be empty!");
        }
        this.primaryKey = primaryKey;
    }

    public List<E> getEntityRecords() {
        return entityRecords;
    }

    public void setEntityRecords(List<E> entityRecords) {
        if (entityRecords != null && !entityRecords.isEmpty()) {
            for (E record : entityRecords) {
                this.entityRecords.add(record);
            }
        }
    }

    public void addEntityRecord(E entity) {
        if (entity == null) {
            return;
        }
        synchronized (entityRecords) {
            if (entityRecords == null) {
                entityRecords = new ArrayList<E>();
            }
            this.entityRecords.add(entity);
        }
    }


    public XLSEntityToRowMapper getEntityToRowMapper() {
        return entityToRowMapper;
    }

    public void setEntityToRowMapper(XLSEntityToRowMapper entityToRowMapper) {
        this.entityToRowMapper = entityToRowMapper;
    }

    public List<XLSColumnDefinition> getColumnsDefinition() {
        return columnsDefinition;
    }

    public void setColumnsDefinition(List<XLSColumnDefinition> columnsDefinition) {
        this.columnsDefinition = columnsDefinition;
    }

    public CellStyle getHeaderColumnsStyle() {
        return headerColumnsStyle;
    }

    public void setHeaderColumnsStyle(CellStyle headerColumnsStyle) {
        this.headerColumnsStyle = headerColumnsStyle;
    }

    public void setColumnStyle(Integer colNo, CellStyle style) {
        colsStyle.put(colNo, style);
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public boolean isRightToLeft() {
        return rightToLeft;
    }

    public void setRightToLeft(boolean rightToLeft) {
        this.rightToLeft = rightToLeft;
    }

    public XLSRowCustomizer getRowCustomizer() {
        return rowCustomizer;
    }

    public void setRowCustomizer(XLSRowCustomizer rowCustomizer) {
        this.rowCustomizer = rowCustomizer;
    }

    public int getLastSheetRecordIndex() {
        return lastSheetRecordIndex;
    }

    public int getAndIncrementLastRowIndex() {
        int result;
        synchronized (lastSheetRecordIndex) {
            result = lastSheetRecordIndex;
            lastSheetRecordIndex++;
        }
        return result;
    }

    public String getReportTitle() {
        if (reportTitle == null) {
            reportTitle = NO_TITLE;
        }
        return reportTitle;
    }

    public void setReportTitle(String reportTitle) {
        this.reportTitle = reportTitle;
    }

    public List<XLSReportSection> getReportSections() {
        if (reportSections == null) {
            reportSections = new ArrayList<XLSReportSection>();
        }
        return reportSections;
    }

    public void setReportSections(List<XLSReportSection> reportSections) {
        this.reportSections = reportSections;
    }

    public HSSFSheet getRealSheet() {
        if (realSheet == null) {
            realSheet = getSheetInstance();
        }
        return realSheet;
    }

    public void setRealSheet(HSSFSheet realSheet) {
        this.realSheet = realSheet;
    }

    public ExcelReportGenerator getReportGenerator() {
        return reportGenerator;
    }

    public void setReportGenerator(ExcelReportGenerator reportGenerator) {
        this.reportGenerator = reportGenerator;
    }

    public List<XLSColorDescription> getColorsDescription() {
        if (colorsDescription == null) {
            colorsDescription = new ArrayList<XLSColorDescription>();
        }
        return colorsDescription;
    }

    public void setColorsDescription(List<XLSColorDescription> colorsDescription) {
        if (colorsDescription == null) {
            throw new IllegalArgumentException("Invalid null XLSColorDescription list!");
        }
        if (colorsDescription.size() > XLSConstants.MAX_COLORS_PER_SHEET) {
            throw new IllegalArgumentException("Invalid colorsDescription (size must be <= 12)");
        }
        this.colorsDescription = colorsDescription;
    }

    public void initColorsDescriptionList() {
        List<XLSColorDescription> colorsDescription = getColorsDescription();
        if (colorsDescription.size() > XLSConstants.MAX_COLORS_PER_SHEET) {
            throw new IllegalArgumentException("Invalid colorsDescription (size must be <= 12)");
        }
        HSSFPalette palette = getWorkbook().getCustomPalette();
        // add color for decussate rows
        addDefaultColors(palette);
        int colorPaletteStartIndex = getSheetNumber() * XLSConstants.MAX_COLORS_PER_SHEET;
        int realIndex = 0;
        for (int colorIndex = 0; colorIndex < colorsDescription.size(); colorIndex++) {
            XLSColorDescription colorDescription = colorsDescription.get(colorIndex);
            if (!colorDescription.isDefaultColor()) {
                colorDescription.setRealIndex(realIndex);
                addToWorkBookPalette(colorDescription, palette, colorPaletteStartIndex);
                realIndex++;
            }
            getColorsDescriptionMap().put(colorDescription.getKey(), colorDescription);
        }
        this.colorsDescription = colorsDescription;
    }

    private void addDefaultColors(HSSFPalette palette) {
        palette.setColorAtIndex(XLSConstants.DEFAULT_ODD_ROW_COLOR_INDEX, (byte) 150, (byte) 255, (byte) 200);
        palette.setColorAtIndex(XLSConstants.DEFAULT_EVEN_ROW_COLOR_INDEX, (byte) 204, (byte) 255, (byte) 204);
        XLSColorDescription oddColorDescription = new XLSColorDescription(XLSConstants.DEFAULT_ODD_ROW_COLOR_KEY, 150, 255, 200, XLSConstants.DEFAULT_ODD_ROW_COLOR_DESCRIPTION);
        oddColorDescription.setColorIndex(XLSConstants.DEFAULT_ODD_ROW_COLOR_INDEX);
        oddColorDescription.setDefaultColor(true);
        XLSColorDescription evenColorDescription = new XLSColorDescription(XLSConstants.DEFAULT_EVEN_ROW_COLOR_KEY, 204, 255, 204, XLSConstants.DEFAULT_EVEN_ROW_COLOR_DESCRIPTION);
        evenColorDescription.setColorIndex(XLSConstants.DEFAULT_EVEN_ROW_COLOR_INDEX);
        evenColorDescription.setDefaultColor(true);
        XLSColorDescription fontColorDescription = new XLSColorDescription(XLSConstants.DEFAULT_FONT_COLOR_KEY, 0, 0, 0, XLSConstants.DEFAULT_FONT_COLOR_DESCRIPTION);
        fontColorDescription.setColorIndex((short) 8);
        fontColorDescription.setDefaultColor(true);
        getColorsDescription().add(oddColorDescription);
        getColorsDescription().add(evenColorDescription);
        getColorsDescription().add(fontColorDescription);
    }

    private void addToWorkBookPalette(XLSColorDescription colorDescription, HSSFPalette palette, int colorPaletteStartIndex) {
        //replacing the standard colors
        int colorRealIndex = XLSConstants.DEFAULT_EVEN_ROW_COLOR_INDEX + 1 + colorDescription.getRealIndex() + colorPaletteStartIndex;
        palette.setColorAtIndex((short) (colorRealIndex),
                (byte) colorDescription.getRed(),  //RGB red (0-255)
                (byte) colorDescription.getGreen(),    //RGB green (0-255)
                (byte) colorDescription.getBlue()     //RGB blue (0-255)
        );
        colorDescription.setColorIndex((short) colorRealIndex);
    }

    public void addColorToSheet(XLSColorDescription xlsColorDescription) {
        if (xlsColorDescription == null) {
            throw new IllegalArgumentException("Invalid null XLSColorDescription!");
        }
        List<XLSColorDescription> colorsDescription = getColorsDescription();
        if (colorsDescription.size() == XLSConstants.MAX_COLORS_PER_SHEET) {
            throw new IllegalArgumentException("Already " + XLSConstants.MAX_COLORS_PER_SHEET + " Colors assigned in this sheet!");
        }
        for (XLSColorDescription description : colorsDescription) {
            if (description.getKey().equals(xlsColorDescription.getKey())) {
                throw new IllegalArgumentException("Color with key " + description.getKey() + "already assigned in this sheet!");
            }
        }
        getColorsDescription().add(xlsColorDescription);
    }


    public void addDummyRecords(int dummyRowsCount) {
        if (dummyRowsCount < 0) {
            throw new IllegalArgumentException("Dummy row count must be greater than 0");
        }
        this.dummyRowsCount = dummyRowsCount;
    }

    public int getDummyRowsCount() {
        return dummyRowsCount;
    }

    public void addDummyEntity(E dummyEntity) {
        if (dummyEntity == null) {
            throw new IllegalArgumentException("dummy entity must not be null!");
        }
        this.dummyEntity = dummyEntity;
    }

    public E getDummyEntity() {
        return dummyEntity;
    }

    public boolean isParsed() {
        return parsed;
    }

    public void setParsed(boolean parsed) {
        this.parsed = parsed;
    }

    public void addBadFormatRecord(int rowIndex, int dummyRowIndex) {
        HSSFRow row = getRow(dummyRowIndex);
        Map<Integer, List<HSSFRow>> badFormatRecords = getBadFormatRecords();
        List<HSSFRow> hssfRows = badFormatRecords.get(rowIndex);
        if (hssfRows == null) {
            hssfRows = new ArrayList<HSSFRow>();
        }
        hssfRows.add(row);
        getBadFormatRecords().put(rowIndex, hssfRows);
    }

    public Map<Integer, List<HSSFRow>> getBadFormatRecords() {
        if (badFormatRecords == null) {
            badFormatRecords = new HashMap<Integer, List<HSSFRow>>();
        }
        return badFormatRecords;
    }


    public int getMainColumnsCount() {
        int firstRowColCount = 0;
        if (isHeaderSet()) {
            for (XLSColumnDefinition definition : getHeaderColumns()) {
                if (!definition.isHidden()) {
                    if (!definition.isRealColumn() || (definition.isRealColumn() && !hasSubColumns(definition))) {
                        firstRowColCount++;
                    }
                }
            }
        }
        return firstRowColCount;
    }

    public List<String> getUniqueColumnKey() {
        if (!uniqueColumnKeyParesed) {
            uniqueKeys = new ArrayList<String>();
            List<XLSColumnDefinition> headerColumns = getHeaderColumns();
            for (XLSColumnDefinition definition : headerColumns) {
                if (definition.isUniqueColumn()) {
                    uniqueKeys.add(definition.getName());
                }
                for (XLSColumnDefinition subDefinition : definition.getSubColumns()) {
                    if (subDefinition.isUniqueColumn()) {
                        uniqueKeys.add(subDefinition.getName());
                    }
                }
            }
            uniqueColumnKeySet = !uniqueKeys.isEmpty();
        }
        return uniqueKeys;
    }

    public boolean isUniqueColumnKeySet() {
        if (!uniqueColumnKeyParesed) {
            getUniqueColumnKey();
        }
        return uniqueColumnKeySet;
    }

    public void setUniqueColumnKeySet(boolean uniqueColumnKeySet) {
        this.uniqueColumnKeySet = uniqueColumnKeySet;
    }

    public List<XLSReportField> getReportInformationFields() {
        if (reportInformationFields == null) {
            reportInformationFields = new ArrayList<XLSReportField>();
        }
        return reportInformationFields;
    }

    public void setReportInformationFields(List<XLSReportField> reportInformationFields) {
        if (reportInformationFields != null && !reportInformationFields.isEmpty()) {
            for (XLSReportField reportField : reportInformationFields) {
                if (reportField.getWidth() <= 0) {
                    throw new RuntimeException("Invalid Report field width value(" + reportField.getWidth() + ")!");
                }
            }
            this.reportInformationFields = reportInformationFields;
        }
    }


    public void addBusinessVariable(String variableName, Object variableValue) {
        if (StringUtils.isEmpty(variableName)) {
            throw new IllegalArgumentException("Invalid business variable name!");
        }
        getBusinessVariables().put(variableName, variableValue);
    }

    public Object getBusinessVariable(String variableName) {
        if (StringUtils.isEmpty(variableName)) {
            throw new IllegalArgumentException("Invalid business variable name!");
        }
        return getBusinessVariables().get(variableName);
    }

    public Map<String, Object> getBusinessVariables() {
        if (businessVariables == null) {
            businessVariables = new HashMap<String, Object>();
        }
        return businessVariables;
    }

    public Map<Integer, List<HSSFRow>> getRawPOIRecords() {
        if (rawPOIRecords == null) {
            rawPOIRecords = new HashMap<Integer, List<HSSFRow>>();
        }
        return rawPOIRecords;
    }

    public void setRawPOIRecords(Map<Integer, List<HSSFRow>> rawPOIRecords) {
        this.rawPOIRecords = rawPOIRecords;
    }

    public String getEmptyRecordsMessage() {
        return emptyRecordsMessage;
    }

    public void setEmptyRecordsMessage(String emptyRecordsMessage) {
        this.emptyRecordsMessage = emptyRecordsMessage;
    }

    public XLSColorDescription getColorDescription(String colorKey) {
        if (StringUtils.isEmpty(colorKey)) {
            throw new IllegalArgumentException("Color key is empty!");
        }
        XLSColorDescription colorDescription = getColorsDescriptionMap().get(colorKey);
        if (colorDescription == null) {
            throw new IllegalArgumentException("Color with key " + colorKey + " does not exist in sheet context!");
        }
        return colorDescription;
    }

    public Map<String, XLSColorDescription> getColorsDescriptionMap() {
        if (colorsDescriptionMap == null) {
            colorsDescriptionMap = new HashMap<String, XLSColorDescription>();
        }
        return colorsDescriptionMap;
    }

    public short getDefaultEvenRowColor() {
        return getColorDescription(XLSConstants.DEFAULT_EVEN_ROW_COLOR_KEY).getColorIndex();
    }

    public short getDefaultOddRowColor() {
        return getColorDescription(XLSConstants.DEFAULT_ODD_ROW_COLOR_KEY).getColorIndex();
    }

    public CellStyle createPOIRowStyle(XLSCellStyle xlsCellStyle) {
        HSSFCellStyle cellStyle;
        if (!getCellStyleMap().containsKey(xlsCellStyle)) {
            cellStyle = getWorkbook().createCellStyle();
            xlsCellStyle.setRealCellStyle(cellStyle);
            getCellStyleMap().put(xlsCellStyle, xlsCellStyle);
        } else {
            cellStyle = getCellStyleMap().get(xlsCellStyle).getRealCellStyle();
        }
        return cellStyle;
    }

    public Map<XLSCellStyle, XLSCellStyle> getCellStyleMap() {
        if (cellStyleMap == null) {
            cellStyleMap = new HashMap<XLSCellStyle, XLSCellStyle>();
        }
        return cellStyleMap;
    }

    public Integer getProcessedEntityCount() {
        return processedEntityCount;
    }

    public void incProcessedEntityCount() {
        synchronized (processedEntityCount) {
            processedEntityCount++;
        }
    }

    public void createCell(String cellValue, int rowNo, int colNo, int width, CellStyle cellStyle, int cellType) {
        HSSFSheet realSheet = getRealSheet();
        HSSFRow row = realSheet.getRow(rowNo);
        if (row == null) {
            row = realSheet.createRow(rowNo);
        }
        HSSFCell cell = row.getCell(colNo);
        if (cell == null) {
            cell = row.createCell(colNo, cellType);
        }
        cell.setCellValue(cellValue);
        cell.setCellStyle(cellStyle);
        for (int index = colNo + 1; index < colNo + width; index++) {
            cell = row.getCell(index);
            if (cell == null) {
                cell = row.createCell(index, cellType);
            }
            cell.setCellValue("");
            cell.setCellStyle(cellStyle);
        }
        this.realSheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo, colNo, colNo + width - 1));
    }

    public boolean hasColorDescriptionTable() {
        List<XLSColorDescription> colorsDescription = getColorsDescription();
        for (XLSColorDescription color : colorsDescription) {
            if (!color.isDefaultColor()) {
                return true;
            }
        }
        return false;
    }

    public String getReadOnlyPassword() {
        return readOnlyPassword;
    }

    public void setReadOnlyPassword(String readOnlyPassword) {
        this.readOnlyPassword = readOnlyPassword;
    }
}
