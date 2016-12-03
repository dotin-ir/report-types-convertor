package ir.dotin.utils.xls.reader;

import ir.dotin.utils.xls.checker.XLSPrimaryKeyChecker;
import ir.dotin.utils.xls.checker.XLSUniqueChecker;
import ir.dotin.utils.xls.domain.XLSColumnDefinition;
import ir.dotin.utils.xls.domain.XLSRecord;
import ir.dotin.utils.xls.domain.XLSSheetContext;
import ir.dotin.utils.xls.mapper.XLSRowToEntityMapper;
import ir.dotin.utils.xls.validator.XLSDocumentValidator;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by r.rastakfard on 6/21/2016.
 */
public class XLSListReader {


    private HSSFWorkbook workbook;
    private HashMap<Integer, XLSSheetContext> sheetsContext = new HashMap<Integer, XLSSheetContext>();


    public XLSListReader(InputStream workbookStream) throws IOException {
        workbook = new HSSFWorkbook(workbookStream);
    }

    public XLSListReader(HSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    public List<CellRangeAddress> getMergedRegions(Sheet sheet, int row) {
        List<CellRangeAddress> cellRangeAddresses = new ArrayList<CellRangeAddress>();
        for (int i = 0; i < sheet.getNumMergedRegions(); ++i) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.getFirstRow() <= row && range.getLastRow() >= row) {
                cellRangeAddresses.add(range);
            }
        }
        return cellRangeAddresses;
    }


    public void setDocumentValidator(Integer sheetNo, XLSDocumentValidator documentValidator) {
        XLSSheetContext context = getSheetContext(sheetNo);
        context.setDocumentValidator(documentValidator);
    }

    public XLSSheetContext getSheetContext(Integer sheetNo) {
        XLSSheetContext context = sheetsContext.get(sheetNo);
        if (context == null) {
            context = new XLSSheetContext(workbook, sheetNo);
            sheetsContext.put(sheetNo, context);
        }
        return context;
    }

    public void parseSheet(int sheetNo) throws Exception {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        parseSheetRecords(sheetContext);
    }

    public void parseAllSheet() throws Exception {
        for (int sheetNo : sheetsContext.keySet()) {
            XLSSheetContext xlsSheetContext = sheetsContext.get(sheetNo);
            parseSheetRecords(xlsSheetContext);
        }
    }

    private void parseSheetRecords(XLSSheetContext sheetContext) throws Exception {
        if (!sheetContext.isParsed()) {
            parseAndOrderByRealHeaderColumns(sheetContext);
            sheetContext.validateDocumentSheet();
            HSSFSheet realSheet = sheetContext.getRealSheet();
            int firstRowIndex = 0;
            if (sheetContext.isHeaderSet()) {
                firstRowIndex = sheetContext.getFirstRowIndex();
            }
            int realRecordCount = 0;
            for (int rowIndex = firstRowIndex; rowIndex <= realSheet.getLastRowNum(); ) {
                HSSFRow row = realSheet.getRow(rowIndex);
                if (row == null) continue;
                int recordRealRowSize = getRecordRealRowSize(row, rowIndex, sheetContext);
                processRow(sheetContext, row, rowIndex, recordRealRowSize, realRecordCount + 1);
                rowIndex += recordRealRowSize;
                realRecordCount++;
            }

            List<XLSUniqueChecker> uniqueCheckers = sheetContext.getUniqueCheckers();
            if (!sheetContext.isUniqueCheckerStatus()) {
                List<XLSRecord> rawRecords = sheetContext.getRawRecords();
                for (XLSRecord record : rawRecords) {
                    for (XLSUniqueChecker uniqueChecker : uniqueCheckers) {
                        if (!uniqueChecker.getStatus()) {
                            String hash = uniqueChecker.getHash(record, sheetContext);
                            if (sheetContext.getErrorRecordsWithKey().contains(hash)) {
                                sheetContext.getErrorRecords().add(record);
                            }
                        }
                    }
                }
                sheetContext.getRawRecords().removeAll(sheetContext.getErrorRecords());
            }
            sheetContext.createResultRecordsObject();
            sheetContext.setParsed(true);
        }
    }

    private int getRecordRealRowSize(HSSFRow row, int rowIndex, XLSSheetContext sheetContext) {
        int maxRowSize = 1;
        List<CellRangeAddress> mergedRegions = getMergedRegions(sheetContext.getRealSheet(), rowIndex);
        for (CellRangeAddress rangeAddress : mergedRegions) {
            int mergedRegionRowSize = getMergedRegionRowSize(rangeAddress);
            if (maxRowSize < mergedRegionRowSize) {
                maxRowSize = mergedRegionRowSize;
            }
        }
        return maxRowSize;
    }

    private int getMergedRegionRowSize(CellRangeAddress rangeAddress) {
        return rangeAddress.getLastRow() - rangeAddress.getFirstRow() + 1;
    }

    private int getMergedRegionColSize(CellRangeAddress rangeAddress) {
        return rangeAddress.getLastColumn() - rangeAddress.getFirstColumn() + 1;
    }

    private void processRow(XLSSheetContext sheetContext, HSSFRow recordFirstRow, int rowIndex, int recordRealRowSize, int realRecordCount) {
        boolean badFormatRecordFound;
        int tmpRecordRealRowSize = recordRealRowSize;
        int nextRecordFirstRowIndex = rowIndex + tmpRecordRealRowSize;
        int tmpRowIndex = rowIndex;
        XLSDocumentValidator validator = sheetContext.getDocumentValidator();
        List<XLSUniqueChecker> uniqueCheckers = sheetContext.getUniqueCheckers();
        Map<String, List<Map<String, String>>> currentRecord = new HashMap<String, List<Map<String, String>>>();
        XLSRecord currentXlsRecord = new XLSRecord();
        boolean validRow;
        badFormatRecordFound = fillCurrentRecord(sheetContext, recordFirstRow, rowIndex, nextRecordFirstRowIndex, tmpRowIndex, currentRecord, currentXlsRecord);
        if (!badFormatRecordFound) {
            if (isEmptyRecord(currentRecord)) {
                return;
            }

            currentXlsRecord.setRecordData(currentRecord);
            currentXlsRecord.setRecordColumns(sheetContext.getHeaderColumns());
            validRow = validator.validateRow(currentXlsRecord, sheetContext);
            boolean isValidRowBasedOnUniqueness = true;
            for (XLSUniqueChecker uniqueChecker : uniqueCheckers) {
                isValidRowBasedOnUniqueness &= uniqueChecker.checkUniqueRow(currentXlsRecord, sheetContext);
            }
            if (validRow && isValidRowBasedOnUniqueness) {
                sheetContext.getRawRecords().add(currentXlsRecord);
            } else {
                if (!isValidRowBasedOnUniqueness) {
                    String hash = null;
                    for (XLSUniqueChecker uniqueChecker : uniqueCheckers) {
                        if (!uniqueChecker.getStatus()) {
                            hash = uniqueChecker.getHash(currentXlsRecord, sheetContext);
                        }
                    }
                    if (hash == null) {
                        throw new RuntimeException("at least getHash method of one uniqueChecker mus be work!");
                    }
                    sheetContext.getErrorRecordsWithKey().add(hash);
                }
                sheetContext.getErrorRecords().add(currentXlsRecord);
            }
        } else {
            int tmpRealRowIndex = rowIndex;
            while (tmpRealRowIndex < nextRecordFirstRowIndex) {
                sheetContext.addBadFormatRecord(realRecordCount, tmpRealRowIndex);
                tmpRealRowIndex++;
            }
        }
    }

    private boolean fillCurrentRecord(XLSSheetContext sheetContext, HSSFRow recordFirstRow, int rowIndex,
                                      int nextRecordFirstRowIndex, int tmpRowIndex, Map<String, List<Map<String, String>>> currentRecord, XLSRecord currentXlsRecord) {
        boolean badFormatRecordFound = false;
        List<XLSColumnDefinition> headerColumns = sheetContext.getHeaderColumns();
        int realCellIndex = 0;
        for (int colIndex = 0; colIndex < headerColumns.size(); colIndex++) {
            XLSColumnDefinition columnDefinition = headerColumns.get(colIndex);
            List<Map<String, String>> resultCellData = new ArrayList<Map<String, String>>();
            if (columnDefinition.isRealColumn()) {
                HSSFRow tmpRecordFirstRow = recordFirstRow;
                try {
                    resultCellData = getCellData(sheetContext, nextRecordFirstRowIndex, tmpRowIndex, columnDefinition, tmpRecordFirstRow, realCellIndex);
                } catch (RuntimeException ex) {
                    badFormatRecordFound = true;
                    return badFormatRecordFound;
                }
                currentRecord.put(columnDefinition.getName(), resultCellData);
                realCellIndex += columnDefinition.getWidth();
            } else {
                HSSFRow tmpRecordFirstRow = recordFirstRow;
                int tmpRealCellIndex = realCellIndex;
                while (tmpRowIndex < nextRecordFirstRowIndex) {
                    int cellRealRowSize = 0;
                    int firstCellRealRowSize = 0;
                    Map<String, String> cellData = new HashMap<String, String>();
                    for (int subColIndex = 0; subColIndex < columnDefinition.getSubColumns().size(); subColIndex++) {
                        List<XLSColumnDefinition> subColumns = columnDefinition.getSubColumns();
                        XLSColumnDefinition subColumn = subColumns.get(subColIndex);
                        HSSFCell tmpCell = tmpRecordFirstRow.getCell(tmpRealCellIndex);
                        cellRealRowSize = getCellRealRowSize(tmpRealCellIndex, sheetContext, tmpCell);
                        try {
                            checkRealColumnWidthWithDefinition(sheetContext, subColumn, realCellIndex, tmpCell);
                        } catch (RuntimeException ex) {
                            badFormatRecordFound = true;
                            return badFormatRecordFound;
                        }
                        if (subColIndex == 0) {
                            firstCellRealRowSize = cellRealRowSize;
                        } else {
                            if (cellRealRowSize != firstCellRealRowSize) {
                                badFormatRecordFound = true;
                                return badFormatRecordFound;
                            }
                        }
                        String cellValue = sheetContext.getCellValue(tmpCell);
                        if (cellValue == null) {
                            cellValue = columnDefinition.getDefaultEmptyData();
                        }
                        cellData.put(subColumn.getName(), cellValue);
                        tmpRealCellIndex += subColumn.getWidth();
                    }
                    realCellIndex = tmpRealCellIndex;
                    tmpRealCellIndex -= columnDefinition.getRealColumnWidth();
                    tmpRowIndex += cellRealRowSize;
                    tmpRecordFirstRow = sheetContext.getRow(tmpRowIndex);
                    resultCellData.add(cellData);
                }
                currentRecord.put(columnDefinition.getName(), resultCellData);
            }
            tmpRowIndex = rowIndex;
        }
        return badFormatRecordFound;
    }

    private List<Map<String, String>> getCellData(XLSSheetContext sheetContext, int nextRecordFirstRowIndex, int tmpRowIndex, XLSColumnDefinition columnDefinition, HSSFRow tmpRecordFirstRow, int realCellIndex) {
        List<Map<String, String>> resultCellData = new ArrayList<Map<String, String>>();
        while (tmpRowIndex < nextRecordFirstRowIndex) {
            HSSFCell tmpCell = tmpRecordFirstRow.getCell(realCellIndex);
            Map<String, String> cellData = new HashMap<String, String>();
            int cellRealRowSize = getCellRealRowSize(realCellIndex, sheetContext, tmpCell);
            checkRealColumnWidthWithDefinition(sheetContext, columnDefinition, realCellIndex, tmpCell);
            String cellValue = sheetContext.getCellValue(tmpCell);
            if (cellValue == null) {
                cellValue = columnDefinition.getDefaultEmptyData();
            }
            cellData.put(columnDefinition.getName(), cellValue);
            tmpRowIndex += cellRealRowSize;
            resultCellData.add(cellData);
            tmpRecordFirstRow = sheetContext.getRow(tmpRowIndex);
        }
        return resultCellData;
    }

    private void checkRealColumnWidthWithDefinition(XLSSheetContext sheetContext, XLSColumnDefinition columnDefinition, int realCellIndex, HSSFCell tmpCell) {
        int cellRealColSize = getCellRealColSize(realCellIndex, sheetContext, tmpCell);
        if (cellRealColSize != columnDefinition.getRealColumnWidth()) {
            throw new RuntimeException("Bad Column width (mismatch with columns definition)!");
        }
    }



    /*private List<Map<String, String>> getCellValue(HSSFCell cell, XLSSheetContext sheetContext, XLSColumnDefinition columnDefinition, int recordRealRowSize) {
        int cellRealRowSize = getCellRealRowSize(cellIndex, cell);
        List<Map<String, String>> result = new ArrayList<Map<String, String>>();
        Map<String, String> cellData = new HashMap<String, String>();
        while (cellRealRowSize <= recordRealRowSize) {
            if (!columnDefinition.isRealColumn()) {
                for (XLSColumnDefinition subColumn : columnDefinition.getSubColumns()) {
                    cellData.put(subColumn.getName(), sheetContext.getCellValue(cell));
                }
            } else {
                cellData.put(columnDefinition.getName(), sheetContext.getCellValue(cell));
            }
            result.add(cellData);
            cellRealRowSize++;
        }
        return null;
    }*/

    private int getCellRealRowSize(int cellIndex, XLSSheetContext sheetContext, HSSFCell cell) {
        int maxRowSize = 1;
        if (cell == null) return maxRowSize;
        List<CellRangeAddress> mergedRegions = getMergedRegions(sheetContext.getRealSheet(), cell.getRowIndex());
        for (CellRangeAddress rangeAddress : mergedRegions) {
            if (rangeAddress.getFirstColumn() <= cell.getColumnIndex() && rangeAddress.getLastColumn() >= cell.getColumnIndex()) {
                return getMergedRegionRowSize(rangeAddress);
            }
        }
        return maxRowSize;
    }

    private int getCellRealColSize(int realCellIndex, XLSSheetContext sheetContext, HSSFCell cell) {
        int maxColSize = 1;
        if (cell == null) return maxColSize;
        List<CellRangeAddress> mergedRegions = getMergedRegions(sheetContext.getRealSheet(), cell.getRowIndex());
        for (CellRangeAddress rangeAddress : mergedRegions) {
            if (rangeAddress.getFirstColumn() <= cell.getColumnIndex() && rangeAddress.getLastColumn() >= cell.getColumnIndex()) {
                return getMergedRegionColSize(rangeAddress);
            }
        }
        return maxColSize;
    }

    private boolean isEmptyRecord(Map<String, List<Map<String, String>>> record) {
        for (String key : record.keySet()) {
            List<Map<String, String>> map = record.get(key);
            for (Map<String, String> values : map) {
                for (String realKey : values.keySet()) {
                    if (!StringUtils.isEmpty(values.get(realKey))) {
                        return false;
                    }
                }
            }
        }
        return true;
    }

    public void setHeaderColumns(int sheetNo, List<XLSColumnDefinition> headerColumns) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.setHeaderColumns(headerColumns);
    }

    public boolean hasColWithHeaderName(int sheetNo, String colName) {
        return getSheetContext(sheetNo).hasColWithHeaderName(colName);
    }

    public void addUniqueChecker(int sheetNo, XLSUniqueChecker uniqueChecker) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.addUniqueChecker(uniqueChecker);
    }

    public void setResultMapper(int sheetNo, XLSRowToEntityMapper resultMapper) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.setResultMapper(resultMapper);
    }

    public void setPrimaryKey(int sheetNo, String primaryKey) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.addUniqueChecker(new XLSPrimaryKeyChecker(primaryKey));
        sheetContext.setPrimaryKey(primaryKey);
    }

    public void parseAndOrderByRealHeaderColumns(XLSSheetContext sheetContext) throws Exception {
        List<XLSColumnDefinition> developerHeaderColumns = sheetContext.getHeaderColumns();
        List<XLSColumnDefinition> realHeaderColumns = new ArrayList<XLSColumnDefinition>();
        int headerRowCount = 1;
        for (XLSColumnDefinition definition : developerHeaderColumns) {
            if (sheetContext.hasSubColumns(definition)) {
                headerRowCount = 2;
                break;
            }
        }
        HSSFRow firstRow = sheetContext.getRow(0);
        for (int realHeaderColIndex = 0; realHeaderColIndex < firstRow.getLastCellNum(); ) {
            HSSFCell cell = firstRow.getCell(realHeaderColIndex);
            if (cell == null) continue;
            int width = getCellRealColSize(realHeaderColIndex, sheetContext, cell);
            String headerColValue = sheetContext.getCellValue(cell);
            if (headerColValue == null) {
                headerColValue = "";
            }
            boolean found = false;
            for (XLSColumnDefinition definition : developerHeaderColumns) {
                if (!definition.isHidden()) {
                    if (!definition.isRealColumn() || (definition.isRealColumn() && !sheetContext.hasSubColumns(definition))) {
                        if (definition.getfName().equals(headerColValue)) {
                            if (width != definition.getRealColumnWidth()) {
                                throw new Exception("Header with column name [ " + headerColValue + " ] mismatch width in definition!");
                            }
                            realHeaderColumns.add(definition);
                            found = true;
                            break;
                        }
                    }
                }
            }
            if (!found) {
                throw new Exception("Header with column name [ " + headerColValue + " ] not found in definition!");
            }
            realHeaderColIndex += width;
        }

        checkMandatoryColumns(developerHeaderColumns, realHeaderColumns);

        if (headerRowCount == 2) {
            int secondRowRealColCount = 0;
            HSSFRow subColsRow = sheetContext.getRow(1);
            boolean found = false;
            for (XLSColumnDefinition headerCol : realHeaderColumns) {
                List<XLSColumnDefinition> realSubCols = new ArrayList<XLSColumnDefinition>();
                for (int realHeaderColIndex = 0; realHeaderColIndex < subColsRow.getLastCellNum(); ) {
                    HSSFCell cell = subColsRow.getCell(realHeaderColIndex);
                    int width = getCellRealColSize(realHeaderColIndex, sheetContext, cell);
                    String subColValue = sheetContext.getCellValue(cell);
                    if (subColValue == null) {
                        subColValue = "";
                    }
                    List<XLSColumnDefinition> subColumns = headerCol.getSubColumns();
                    for (XLSColumnDefinition subDefinition : subColumns) {
                        if (!subDefinition.isHidden()) {
                            if (subDefinition.getfName().equals(subColValue)) {
                                if (width != subDefinition.getWidth()) {
                                    throw new Exception("SubHeader with column name [ " + subColValue + " ] mismatch width in definition!");
                                }
                                realSubCols.add(subDefinition);
                                secondRowRealColCount++;
                                found = true;
                                break;
                            }
                            if (!found) {
                                throw new Exception("SubHeader with column name [ " + subColValue + " ] not found in definition!");
                            }
                        }
                    }
                    realHeaderColIndex += width;
                }

                checkMandatoryColumns(headerCol.getSubColumns(), realSubCols);

                if (headerCol.getSubColumns().size() != secondRowRealColCount) {
                    throw new Exception("Header sub columns for header column [ " + headerCol.getfName() + " ]mismatch with definition!");
                }

                headerCol.setSubColumns(realSubCols);
                secondRowRealColCount = 0;
            }

        }
        sheetContext.setHeaderColumns(realHeaderColumns);
    }

    private void checkMandatoryColumns(List<XLSColumnDefinition> developerHeaderColumns, List<XLSColumnDefinition> realHeaderColumns) {
        for (XLSColumnDefinition definition : developerHeaderColumns) {
            boolean found = false;
            for (XLSColumnDefinition realDefinition : realHeaderColumns) {
                if (realDefinition.getName().equals(definition.getName())) {
                    found = true;
                    break;
                }
            }
            if (!found) {
                if (definition.isMandatory()) {
                    throw new RuntimeException("Header column with name : " + definition.getName() + " not exists in file!");
                }
            }
        }
    }

    public void addBusinessVariable(int sheetNo, String variableName, Object variableValue) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.addBusinessVariable(variableName, variableValue);
    }
}
