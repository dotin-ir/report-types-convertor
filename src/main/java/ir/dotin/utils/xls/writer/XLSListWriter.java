package ir.dotin.utils.xls.writer;

import ir.dotin.utils.xls.checker.XLSUtils;
import ir.dotin.utils.xls.clone.DeepCopy;
import ir.dotin.utils.xls.domain.*;
import ir.dotin.utils.xls.mapper.XLSEntityToRowMapper;
import ir.dotin.utils.xls.renderer.XLSDefaultRowCustomizer;
import ir.dotin.utils.xls.renderer.XLSRowCustomizer;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * Created by r.rastakfard on 7/2/2016.
 */
public class XLSListWriter<E> extends XLSBaseWriter {


    private void setRowCustomizer(int sheetNo, XLSRowCustomizer rowCustomizer) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.setRowCustomizer(rowCustomizer);
    }

    public void addRawRecord(int sheetNo, HashMap<String, String> record) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.addRawRecord(record);
    }

    public void addRecord(int sheetNo, E entity) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.addEntityRecord(entity);
    }

    public void setColumnStylesheet(int sheetNo, int colNo, CellStyle cellStyle) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.setColumnStyle(colNo, cellStyle);
    }

    private void setFirstRowStylesheet(int sheetNo, CellStyle cellStyle) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.setHeaderColumnsStyle(cellStyle);
    }

    public void setRightToLeft(int sheetNo, boolean rightToLeft) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.setRightToLeft(rightToLeft);
    }

    public void setColumnsDefinition(int sheetNo, List<XLSColumnDefinition> columnsDefinition) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.setColumnsDefinition(columnsDefinition);
    }

    public ByteArrayOutputStream createDocument() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        int sheetIndex = 0;
        for (XLSSheetContext sheetContext : getSheetContextsValues()) {
            sheetIndex += 1;
            createSheet(workbook, sheetIndex, sheetContext);
            List<XLSColumnDefinition> columnsDefinition = sheetContext.getColumnsDefinition();
            createHeaderRow(sheetContext, columnsDefinition);
            createXLSRecords(sheetContext);
        }
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        workbook.write(byteArrayOutputStream);
        byteArrayOutputStream.flush();
        byteArrayOutputStream.close();
        return byteArrayOutputStream;
    }

    public void createReportColorsDescriptionTable(XLSSheetContext sheetContext) {

    }


    private void createXLSRecords(XLSSheetContext sheetContext) {
        XLSEntityToRowMapper entityToRowMapper = sheetContext.getEntityToRowMapper();
        List entityRecords = sheetContext.getEntityRecords();
        List rawRecords = sheetContext.getRawRecords();
        addDummyRecords(sheetContext, entityRecords, rawRecords, entityToRowMapper);
        List<XLSColumnDefinition> columnsDefinition = sheetContext.getColumnsDefinition();
        parseRecords(sheetContext, columnsDefinition, entityToRowMapper, sheetContext.getRowCustomizer());
    }

    protected void addDummyRecords(XLSSheetContext sheetContext, List entityRecords, List rawRecords, XLSEntityToRowMapper entityToRowMapper) {
        if (entityToRowMapper == null) {
            addDummyRawRecordsToSheetRows(sheetContext, rawRecords);
            sheetContext.setRawXLSRecords((List<XLSRecord>) DeepCopy.copy(rawRecords));
        } else {
            addDummyEntityRecordsToSheetRows(sheetContext, entityRecords);
            sheetContext.setEntityRecords((List) DeepCopy.copy(entityRecords));
        }
    }

    private void addDummyEntityRecordsToSheetRows(XLSSheetContext sheetContext, List<E> entityRecords) {
        for (int dummyRowIndex = 0; dummyRowIndex < sheetContext.getDummyRowsCount(); dummyRowIndex++) {
            if (sheetContext.getDummyEntity() == null) {
                throw new IllegalArgumentException("DummyEntity not set on sheet " + sheetContext.getSheetName());
            }
            entityRecords.add((E) sheetContext.getDummyEntity());
        }
    }

    protected void parseRecords(XLSSheetContext sheetContext, List<XLSColumnDefinition> columnsDefinition,
                                XLSEntityToRowMapper entityToRowMapper, XLSRowCustomizer rowCustomizer) {
        parseEntityOrRawRecords(sheetContext, columnsDefinition, entityToRowMapper, rowCustomizer, true);
        parseEntityOrRawRecords(sheetContext, columnsDefinition, entityToRowMapper, rowCustomizer, false);
        parseRealPOIRecords(sheetContext);
        addEmptyRecordsRow(sheetContext, columnsDefinition);

    }

    private void addEmptyRecordsRow(XLSSheetContext sheetContext, List<XLSColumnDefinition> columnsDefinition) {
        if (isEmptySheet(sheetContext)) {
            addEmptyRowWithMessage(sheetContext, columnsDefinition);
        }
    }

    private void addEmptyRowWithMessage(XLSSheetContext sheetContext, List<XLSColumnDefinition> columnsDefinition) {
        int totalWidth = 0;
        for (XLSColumnDefinition definition : columnsDefinition) {
            totalWidth += definition.getRealColumnWidth();
        }
        HSSFSheet realSheet = sheetContext.getRealSheet();
        HSSFRow emptyRow = realSheet.createRow(sheetContext.getAndIncrementLastRowIndex());
        HSSFCell firstCell = emptyRow.createCell(0, Cell.CELL_TYPE_STRING);
        firstCell.setCellValue(sheetContext.getEmptyRecordsMessage());
        HSSFCellStyle emptyRecordsCellStyle = realSheet.getWorkbook().createCellStyle();
        firstCell.setCellStyle(createPOIDataRowStyle(sheetContext, false, true, true, true, true, sheetContext.getDefaultEvenRowColor(), emptyRecordsCellStyle));
        realSheet.addMergedRegion(new CellRangeAddress(sheetContext.getLastSheetRecordIndex() - 1,
                sheetContext.getLastSheetRecordIndex() - 1, 0, totalWidth - 1));
        for (int emptyCellIndex = 1; emptyCellIndex <= totalWidth - 1; emptyCellIndex++) {
            HSSFCell emptyCell = emptyRow.createCell(emptyCellIndex, Cell.CELL_TYPE_STRING);
            emptyCell.setCellValue("");
            emptyCell.setCellStyle(createPOIDataRowStyle(sheetContext, false, true, true, true, true, sheetContext.getDefaultEvenRowColor(), emptyRecordsCellStyle));
        }
    }

    private boolean isEmptySheet(XLSSheetContext sheetContext) {
        return sheetContext.getEntityRecords().isEmpty() && sheetContext.getRawPOIRecords().isEmpty() && sheetContext.getRawRecords().isEmpty();
    }

    private void parseRealPOIRecords(XLSSheetContext sheetContext) {
        Map<Integer, List<HSSFRow>> rawPOIRecords = new HashMap<Integer, List<HSSFRow>>();// TODO sheetContext.getRawPOIRecords();
        if (!rawPOIRecords.isEmpty()) {
            HSSFRow row = sheetContext.getRealSheet().createRow(sheetContext.getAndIncrementLastRowIndex());
            HSSFCell cell = row.createCell(0);
            cell.setCellValue(XLSUtils.getProperty(XLSConstants.RECORDS_WITH_BAD_FORMAT));
            cell.setCellType(Cell.CELL_TYPE_STRING);
            for (Integer rowKey : rawPOIRecords.keySet()) {
                List<HSSFRow> hssfRows = rawPOIRecords.get(rowKey);
                int rowIndex = sheetContext.getLastSheetRecordIndex() - 1;
                for (HSSFRow realRow : hssfRows) {
                    HSSFSheet realSheet = sheetContext.getRealSheet();
                    XLSUtils.copyRow(realSheet, rowIndex, realRow);
                    rowIndex++;
                }
            }
        }

    }

    private void parseEntityOrRawRecords(XLSSheetContext sheetContext, List<XLSColumnDefinition> columnsDefinition,
                                         XLSEntityToRowMapper entityToRowMapper, XLSRowCustomizer rowCustomizer, boolean entityOrRaw) {
        List entityRecords = entityOrRaw ? sheetContext.getEntityRecords() : sheetContext.getRawRecords();
        if (!entityRecords.isEmpty()) {
            Map<String, List<Map<String, String>>> rowValue;
            int lastSheetRecordIndex = sheetContext.getLastSheetRecordIndex();
            CellStyle cellStyle = null;
            HSSFSheet realSheet = sheetContext.getRealSheet();
            Integer rowCount;
            for (int rowIndex = 1 + lastSheetRecordIndex; rowIndex <= entityRecords.size() + lastSheetRecordIndex; rowIndex++) {
                Object rowObject = entityRecords.get(rowIndex - 1 - lastSheetRecordIndex);
                Map<String, XLSCellStyle> cellsStyle;
                if (rowCustomizer == null) {
                    rowCustomizer = getDefaultRowCustomizer();
                }
                if (entityOrRaw) {
                    if (entityToRowMapper == null) {
                        throw new IllegalArgumentException("entityToRowMapper must be set on sheet " + sheetContext.getSheetName());
                    }
                    XLSRowCellsData rowDefinition = entityToRowMapper.map(rowObject, sheetContext);
                    rowValue = rowDefinition.getRow();
                    if (rowValue == null) rowValue = new HashMap<String, List<Map<String, String>>>();
                    rowCount = getMaxColumnsDataSize(rowDefinition.getColumnsDataSize());
                    cellsStyle = entityToRowMapper.getRecordDesign(rowObject, sheetContext);
                } else {
                    rowValue = ((XLSRecord) rowObject).getRecordData();
                    if (rowValue == null) rowValue = new HashMap<String, List<Map<String, String>>>();
                    Map<String, Integer> columnsDataSize = new HashMap<String, Integer>();
                    for (String key : rowValue.keySet()) {
                        columnsDataSize.put(key, rowValue.get(key).size());
                    }
                    rowCount = getMaxColumnsDataSize(columnsDataSize);
                    cellsStyle = rowCustomizer.createRecordStyle(rowValue, sheetContext);
                }

                if (cellsStyle == null) {
                    cellsStyle = getDefaultRowCustomizer().createRecordStyle(rowValue, sheetContext);
                }


                List<HSSFRow> rowsForOneRecord = new ArrayList<HSSFRow>();
                for (int realRowIndex = 0; realRowIndex < rowCount; realRowIndex++) {
                    HSSFRow row = realSheet.createRow(sheetContext.getAndIncrementLastRowIndex());
                    row.setHeight((short) XLSConstants.DEFAULT_ROW_SIZE);
                    rowsForOneRecord.add(row);
                }

                int nextRealColIndex = 0;
                int maxRowHeight = XLSConstants.DEFAULT_ROW_SIZE;
                for (int colIndex = 0; colIndex < columnsDefinition.size(); colIndex++) {
                    XLSColumnDefinition columnDefinition = columnsDefinition.get(colIndex);
                    List<XLSColumnDefinition> processColumns;
                    if (!columnDefinition.isHidden()) {

                        List<Map<String, String>> rowValues = rowValue.get(columnDefinition.getName());
                        if (rowValues == null) {
                            rowValues = new ArrayList<Map<String, String>>();
                        }
                        String xlsCellValue;
                        if (hasSubColumns(columnDefinition)) {
                            processColumns = columnDefinition.getSubColumns();
                        } else {
                            processColumns = Arrays.asList(columnDefinition);
                        }

                        if (!rowValues.isEmpty()) {
                            int realSubColIndex;
                            int nextSubRealColIndex = nextRealColIndex;
                            for (int rowValueIndex = 0; rowValueIndex < rowValues.size(); rowValueIndex++) {
                                nextSubRealColIndex = nextRealColIndex;
                                for (XLSColumnDefinition definition : processColumns) {
                                    Map<String, String> rowData = rowValues.get(rowValueIndex);
                                    Integer width = definition.getWidth();
                                    xlsCellValue = rowData.get(definition.getName());
                                    if (StringUtils.isEmpty(xlsCellValue)) {
                                        xlsCellValue = definition.getDefaultEmptyData();
                                    }
                                    realSubColIndex = nextSubRealColIndex;
                                    nextSubRealColIndex = realSubColIndex + width;
                                    HSSFCell cell = rowsForOneRecord.get(rowValueIndex).createCell(realSubColIndex, Cell.CELL_TYPE_STRING);
                                    int rowHeight = computeRowHeight(xlsCellValue, width);
                                    if (rowHeight > rowsForOneRecord.get(rowValueIndex).getHeight()) {
                                        rowsForOneRecord.get(rowValueIndex).setHeight((short) rowHeight);
                                        maxRowHeight = rowHeight;
                                    } else {
                                        if (maxRowHeight < rowsForOneRecord.get(rowValueIndex).getHeight()) {
                                            maxRowHeight = rowsForOneRecord.get(rowValueIndex).getHeight();
                                        }
                                    }
                                    XLSCellStyle xlsCellStyle = cellsStyle.get(definition.getName());
                                    if (xlsCellStyle == null) {
                                        xlsCellStyle = XLSDefaultRowCustomizer.getDefaultCellStyle(sheetContext);
                                    }
                                    cellStyle = sheetContext.createPOIRowStyle(xlsCellStyle);
                                    if (rowValues.size() < rowCount && rowValueIndex == rowValues.size() - 1) {
                                        realSheet.addMergedRegion(new CellRangeAddress(sheetContext.getLastSheetRecordIndex() - rowCount + rowValueIndex,
                                                sheetContext.getLastSheetRecordIndex() - 1, realSubColIndex, realSubColIndex + width - 1));
                                        for (int emptyRowIndex = rowValueIndex + 1; emptyRowIndex < rowCount; emptyRowIndex++) {
                                            for (int emptyCellIndex = realSubColIndex; emptyCellIndex <= realSubColIndex + width - 1; emptyCellIndex++) {
                                                HSSFCell emptyCell = rowsForOneRecord.get(emptyRowIndex).createCell(emptyCellIndex, Cell.CELL_TYPE_STRING);
                                                emptyCell.setCellValue("");
                                                emptyCell.setCellStyle(fillCellStyle(sheetContext, xlsCellStyle, cellStyle));
                                            }
                                        }
                                    } else {
                                        realSheet.addMergedRegion(new CellRangeAddress(sheetContext.getLastSheetRecordIndex() - rowCount + rowValueIndex,
                                                sheetContext.getLastSheetRecordIndex() - rowCount + rowValueIndex, realSubColIndex, realSubColIndex + width - 1));
                                    }
                                    cell.setCellValue(xlsCellValue);
                                    cell.setCellStyle(fillCellStyle(sheetContext, xlsCellStyle, cellStyle));
                                    for (int emptyCellIndex = realSubColIndex + 1; emptyCellIndex <= realSubColIndex + width - 1; emptyCellIndex++) {
                                        HSSFCell emptyCell = rowsForOneRecord.get(rowValueIndex).createCell(emptyCellIndex, Cell.CELL_TYPE_STRING);
                                        emptyCell.setCellValue("");
                                        emptyCell.setCellStyle(fillCellStyle(sheetContext, xlsCellStyle, cellStyle));
                                    }
                                }
                            }
                            nextRealColIndex = nextSubRealColIndex;
                            int sameRowSize = XLSConstants.DEFAULT_ROW_SIZE;
                            if (maxRowHeight > rowsForOneRecord.size() * XLSConstants.DEFAULT_ROW_SIZE) {
                                sameRowSize = maxRowHeight / rowsForOneRecord.size();
                            }
                            for (int xlsRowIndex = 0; xlsRowIndex < rowsForOneRecord.size(); xlsRowIndex++) {
                                rowsForOneRecord.get(xlsRowIndex).setHeight((short) sameRowSize);
                            }

                        } else {
                            // add empty cols
                            cellStyle = sheetContext.createPOIRowStyle(XLSDefaultRowCustomizer.getDefaultCellStyle(sheetContext));
                            int nextSubRealColIndex = nextRealColIndex;
                            nextSubRealColIndex = createEmptyCells(sheetContext, cellStyle, realSheet, rowCount, rowsForOneRecord, processColumns, nextSubRealColIndex);
                            nextRealColIndex = nextSubRealColIndex;
                        }
                    }
                }
                sheetContext.incProcessedEntityCount();
            }
        }
    }

    private void addDummyRawRecordsToSheetRows(XLSSheetContext sheetContext, List<XLSRecord> rawRecords) {

        for (int dummyRowIndex = 0; dummyRowIndex < sheetContext.getDummyRowsCount(); dummyRowIndex++) {
            List<XLSColumnDefinition> columnsDefinition = sheetContext.getColumnsDefinition();
            Map<String, List<Map<String, String>>> rowDateValue = new HashMap<String, List<Map<String, String>>>();
            for (int colIndex = 0; colIndex < columnsDefinition.size(); colIndex++) {
                XLSColumnDefinition columnDefinition = columnsDefinition.get(colIndex);
                List<XLSColumnDefinition> processColumns;
                Map<String, String> map = new HashMap<String, String>();
                if (hasSubColumns(columnDefinition)) {
                    processColumns = columnDefinition.getSubColumns();
                    for (XLSColumnDefinition definition : processColumns) {
                        map.put(definition.getName(), definition.getDefaultEmptyData());
                    }
                    rowDateValue.put(columnDefinition.getName(), Arrays.asList(map));
                } else {
                    map.put(columnDefinition.getName(), columnDefinition.getDefaultEmptyData());
                    rowDateValue.put(columnDefinition.getName(), Arrays.asList(map));
                }
            }
            XLSRecord xlsRecord = new XLSRecord();
            xlsRecord.setRecordData(rowDateValue);
            rawRecords.add(xlsRecord);
        }
    }

    private int createEmptyCells(XLSSheetContext sheetContext, CellStyle cellStyle, HSSFSheet realSheet, Integer rowCount, List<HSSFRow> rowsForOneRecord, List<XLSColumnDefinition> processColumns, int nextSubRealColIndex) {
        int realSubColIndex;
        for (XLSColumnDefinition definition : processColumns) {
            Integer width = definition.getWidth();
            realSubColIndex = nextSubRealColIndex;
            nextSubRealColIndex = realSubColIndex + width;
            HSSFRow emptyRow = rowsForOneRecord.get(0);
            emptyRow.setHeight((short) XLSConstants.DEFAULT_ROW_SIZE);
            HSSFCell cell = emptyRow.createCell(realSubColIndex, Cell.CELL_TYPE_STRING);
            cell.setCellValue("");
            XLSCellStyle xlsCellStyle = XLSDefaultRowCustomizer.getDefaultCellStyle(sheetContext);
            cell.setCellStyle(fillCellStyle(sheetContext, xlsCellStyle, cellStyle));
            realSheet.addMergedRegion(new CellRangeAddress(sheetContext.getLastSheetRecordIndex() - rowCount,
                    sheetContext.getLastSheetRecordIndex() - 1, realSubColIndex, realSubColIndex + width - 1));
            for (int emptyCellIndex = realSubColIndex + 1; emptyCellIndex <= realSubColIndex + width - 1; emptyCellIndex++) {
                HSSFCell emptyCell = emptyRow.createCell(emptyCellIndex, Cell.CELL_TYPE_STRING);
                emptyCell.setCellValue("");
                emptyCell.setCellStyle(fillCellStyle(sheetContext, xlsCellStyle, cellStyle));
            }

            for (int emptyRowIndex = 1; emptyRowIndex < rowCount; emptyRowIndex++) {
                for (int emptyCellIndex = realSubColIndex; emptyCellIndex <= realSubColIndex + width - 1; emptyCellIndex++) {
                    HSSFCell emptyCell = rowsForOneRecord.get(emptyRowIndex).createCell(emptyCellIndex, Cell.CELL_TYPE_STRING);
                    emptyCell.setCellValue("");
                    emptyCell.setCellStyle(fillCellStyle(sheetContext, xlsCellStyle, cellStyle));
                }
            }
        }
        return nextSubRealColIndex;
    }

    private Integer getMaxColumnsDataSize(Map<String, Integer> columnsDataSize) {
        Integer max = 1;
        for (String key : columnsDataSize.keySet()) {
            Integer columnDataSize = columnsDataSize.get(key);
            if (columnDataSize > max) {
                max = columnDataSize;
            }
        }
        return max;
    }

    protected void createHeaderRow(XLSSheetContext sheetContext, List<XLSColumnDefinition> columnsDefinition) {
        HSSFSheet sheet = sheetContext.getRealSheet();
        boolean hasSubColumns = false;
        HSSFRow firstRow = sheet.createRow(sheetContext.getAndIncrementLastRowIndex());
        CellStyle headerStyle = sheetContext.getWorkbook().createCellStyle();
        int realCellIndex;
        int nextRealColIndex = 0;
        Integer totalSubColumnsWidth;
        for (int cellIndex = 0; cellIndex < columnsDefinition.size(); cellIndex++) {
            totalSubColumnsWidth = 0;
            XLSColumnDefinition xlsColumnDefinition = columnsDefinition.get(cellIndex);
            if (!xlsColumnDefinition.isHidden()) {
                if (hasSubColumns(xlsColumnDefinition)) {
                    hasSubColumns = true;
                    totalSubColumnsWidth = xlsColumnDefinition.getTotalSubColumnsWidth();
                }
                Integer width;
                if (totalSubColumnsWidth > 0) {
                    width = totalSubColumnsWidth;
                } else {
                    width = xlsColumnDefinition.getWidth();
                }
                realCellIndex = nextRealColIndex;
                nextRealColIndex = realCellIndex + width;
                headerStyle = createCell(sheetContext, sheet, firstRow, headerStyle, realCellIndex, nextRealColIndex, xlsColumnDefinition, width);
            }
        }

        nextRealColIndex = 0;
        if (hasSubColumns) {
            HSSFRow secondRow = sheet.createRow(sheetContext.getAndIncrementLastRowIndex());
            for (int cellIndex = 0; cellIndex < columnsDefinition.size(); cellIndex++) {
                XLSColumnDefinition xlsColumnDefinition = columnsDefinition.get(cellIndex);
                if (!xlsColumnDefinition.isHidden()) {
                    if (hasSubColumns(xlsColumnDefinition)) {
                        for (int subCellIndex = 0; subCellIndex < xlsColumnDefinition.getSubColumns().size(); subCellIndex++) {
                            XLSColumnDefinition subColDefinition = xlsColumnDefinition.getSubColumns().get(subCellIndex);
                            Integer width = subColDefinition.getWidth();
                            realCellIndex = nextRealColIndex;
                            nextRealColIndex = realCellIndex + width;
                            headerStyle = createCell(sheetContext, sheet, secondRow, headerStyle, realCellIndex, nextRealColIndex, subColDefinition, width);
                        }
                    } else {
                        // merge with above cell
                        Integer width = xlsColumnDefinition.getWidth();
                        realCellIndex = nextRealColIndex;
                        nextRealColIndex = realCellIndex + width;
                        headerStyle = createEmptyCell(sheetContext, sheet, headerStyle, realCellIndex, nextRealColIndex, secondRow, xlsColumnDefinition, width);
                        sheet.addMergedRegion(new CellRangeAddress(sheetContext.getLastSheetRecordIndex() - 2,
                                sheetContext.getLastSheetRecordIndex() - 1, realCellIndex, realCellIndex + width - 1));
                    }
                }
            }
        }


    }

    private CellStyle createEmptyCell(XLSSheetContext sheetContext, HSSFSheet sheet, CellStyle headerStyle, int realCellIndex, int nextRealColIndex, HSSFRow secondRow, XLSColumnDefinition xlsColumnDefinition, Integer width) {
        HSSFCell cell = secondRow.createCell(realCellIndex, Cell.CELL_TYPE_STRING);
        sheet.addMergedRegion(new CellRangeAddress(sheetContext.getLastSheetRecordIndex() - 1,
                sheetContext.getLastSheetRecordIndex() - 1, realCellIndex, realCellIndex + width - 1));
        cell.setCellValue("");
        headerStyle = createPOIHeaderRowStyle(sheetContext, xlsColumnDefinition.getColor(), true, true, true, true, headerStyle);
        cell.setCellStyle(headerStyle);
        for (int emptyCellIndex = realCellIndex + 1; emptyCellIndex <= nextRealColIndex - 1; emptyCellIndex++) {
            HSSFCell emptyCell = secondRow.createCell(emptyCellIndex, Cell.CELL_TYPE_STRING);
            emptyCell.setCellValue("");
            emptyCell.setCellStyle(headerStyle);
        }
        return headerStyle;
    }

    private CellStyle createCell(XLSSheetContext sheetContext, HSSFSheet sheet, HSSFRow row, CellStyle cellStyle,
                                 int realCellIndex, int nextRealColIndex, XLSColumnDefinition xlsColumnDefinition, Integer width) {
        HSSFCell cell = row.createCell(realCellIndex, Cell.CELL_TYPE_STRING);
        sheet.addMergedRegion(new CellRangeAddress(sheetContext.getLastSheetRecordIndex() - 1,
                sheetContext.getLastSheetRecordIndex() - 1, realCellIndex, realCellIndex + width - 1));
        cell.setCellValue(xlsColumnDefinition.getfName());
        cellStyle = createPOIHeaderRowStyle(sheetContext, xlsColumnDefinition.getColor(), true, true, true, true, cellStyle);
        cell.setCellStyle(cellStyle);
        for (int emptyCellIndex = realCellIndex + 1; emptyCellIndex <= nextRealColIndex - 1; emptyCellIndex++) {
            HSSFCell emptyCell = row.createCell(emptyCellIndex, Cell.CELL_TYPE_STRING);
            emptyCell.setCellValue("");
            emptyCell.setCellStyle(cellStyle);
        }
        return cellStyle;
    }

    private boolean hasSubColumns(XLSColumnDefinition xlsColumnDefinition) {
        return xlsColumnDefinition.getSubColumns() != null && !xlsColumnDefinition.getSubColumns().isEmpty();
    }

    public void setRawRecords(int sheetNo, List<Map<String, String>> rawRecords) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.setRawRecords(rawRecords);
    }

    public void addRawEnitytRecords(int sheetNo, Map<String, List<Map<String, String>>> rawRecord) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.addRawEntityRecord(rawRecord);
    }

    public void setRawEnitytRecords(int sheetNo, List<Map<String, List<Map<String, String>>>> rawRecords) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.setRawEntityRecord(rawRecords);
    }

    public void setEntityToRowMapper(int sheetNo, XLSEntityToRowMapper xlsEntityToRowMapper) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.setEntityToRowMapper(xlsEntityToRowMapper);
    }

    public void setHeaderMapping(int sheetNo, List<XLSColumnDefinition> headers) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.setHeaderColumns(headers);
    }


    public void setRecords(int sheetNo, List<E> records) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.setEntityRecords(records);
    }

    public void addDummyRecords(int sheetNo, int dummyRowsCount) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.addDummyRecords(dummyRowsCount);
    }

    public void addDummyRecords(int sheetNo, int dummyRowsCount, E dummyEntity) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.addDummyRecords(dummyRowsCount);
        sheetContext.addDummyEntity(dummyEntity);
    }
}
