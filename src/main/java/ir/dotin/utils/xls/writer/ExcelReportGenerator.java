package ir.dotin.utils.xls.writer;

import ir.dotin.utils.xls.checker.XLSUtils;
import ir.dotin.utils.xls.domain.*;
import ir.dotin.utils.xls.renderer.XLSRowCustomizer;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFRegionUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * Created by r.rastakfard on 7/10/2016.
 */
public class ExcelReportGenerator extends XLSBaseWriter {

    public int REPORT_TITLE_COL_SIZE = 12;
    private List<XLSReportField> conditions;
    private List<XLSColumnDefinition> headerForAllSection;

    public List<XLSReportField> getConditions() {
        if (conditions == null) {
            conditions = new ArrayList<XLSReportField>();
        }
        return conditions;
    }

    public void setConditions(List<XLSReportField> conditions) {
        if (conditions != null) {
            for (XLSReportField field : conditions) {
                if (StringUtils.isEmpty(field.getName())) {
                    throw new RuntimeException("condition label must not be empty!");
                }
            }
        }
        this.conditions = conditions;
    }

    public void setReportSections(int sheetNo, List<XLSReportSection> reportSections) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.setReportSections(reportSections);
    }

    public String getReportTitle(int sheetNo) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        return sheetContext.getReportTitle();
    }

    public void setReportTitle(int sheetNo, String reportTitle) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.setReportTitle(reportTitle);
    }

    public ExcelReportGenerator addReportSection(int sheetNo, XLSListReportSection section) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.getReportSections().add(section);
        return this;
    }


    @Override
    public ByteArrayOutputStream createDocument() throws IOException {
        int sheetIndex = 1;
        Collection<XLSSheetContext> sheetsContextValues = sheetsContext.values();
        refreshBasicInfoBasedOnSheetContexts();
        for (XLSSheetContext sheetContext : sheetsContextValues) {
            sheetContext.setReportGenerator(this);
            createSheet(sheetIndex, sheetContext);
            createReportTitleRow(sheetContext);
            createReportInformationRows(sheetContext);
            createReportConditionSection(sheetContext);
            createReportColorsDescriptionTable(sheetContext);
            List<XLSListReportSection> reportSections = sheetContext.getReportSections();
            checkRowCountLimit(reportSections);
            for (XLSReportSection section : reportSections) {
                section.setWorkBook(getWorkBook());
                section.setBasicInfos(getBasiceInfos());
                section.getRenderer().renderSection(sheetContext, section);
            }
            sheetIndex = generateSheetIndex(sheetContext, sheetIndex);
        }
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        getWorkBook().write(byteArrayOutputStream);
        byteArrayOutputStream.flush();
        byteArrayOutputStream.close();
        return byteArrayOutputStream;
    }

    public void createReportColorsDescriptionTable(XLSSheetContext sheetContext) {
        if (sheetContext.hasColorDescriptionTable()) {
            int colorColWidth = 6;
            sheetContext.createCell(XLSUtils.getProperty(XLSConstants.DEFAULT_COLORS_DESCRIPTION_TABLE_TITLE), 1, REPORT_TITLE_COL_SIZE + 1, 6, createPOIHeaderRowStyle(sheetContext, IndexedColors.LIGHT_ORANGE.getIndex(), true, true, true, true, 12), Cell.CELL_TYPE_STRING);
            List<XLSColorDescription> colorsDescription = sheetContext.getColorsDescription();
            int colorRowIndex = 2;
            int colorColStep = 0;
            for (XLSColorDescription color : colorsDescription) {
                if (!color.isDefaultColor()) {
                    if (colorRowIndex % 5 == 0) {
                        colorColStep += colorColWidth;
                        colorRowIndex = 2;
                    }
                    if (!color.getKey().equals(XLSConstants.DEFAULT_FONT_COLOR_KEY)) {
                        sheetContext.createCell(color.getDescription(), colorRowIndex, REPORT_TITLE_COL_SIZE + 1 + colorColStep, colorColWidth, createPOIHeaderRowStyle(sheetContext, color.getColorIndex(), true, true, true, true, 12), Cell.CELL_TYPE_STRING);
                        colorRowIndex++;
                    }
                }
            }

            for (int emptyIndex = colorRowIndex; emptyIndex < 5; emptyIndex++) {
                sheetContext.createCell("", emptyIndex, REPORT_TITLE_COL_SIZE + 1 + colorColStep, colorColWidth, createPOIHeaderRowStyle(sheetContext, IndexedColors.WHITE.getIndex(), true, true, true, true, 12), Cell.CELL_TYPE_STRING);
            }
            sheetContext.getRealSheet().addMergedRegion(new CellRangeAddress(colorRowIndex, 4, REPORT_TITLE_COL_SIZE + 1 + colorColStep, REPORT_TITLE_COL_SIZE + 1 + colorColStep + colorColWidth - 1));
            System.out.println();
        }
    }

    protected List<XLSColumnDefinition> getSheetContextColumnsDefinition(XLSSheetContext sheetContext) {
        List<XLSColumnDefinition> result = new ArrayList<XLSColumnDefinition>();
        List<XLSReportSection> reportSections = sheetContext.getReportSections();
        for (XLSReportSection section : reportSections) {
            result.addAll(section.getHeaderCols());
        }

        return result;
    }

    private void checkRowCountLimit(List<XLSListReportSection> reportSections) {
        Integer totalRecords = 0;
        for (XLSReportSection section : reportSections) {
            totalRecords += section.getRecordsCount();
        }
        if (totalRecords > 0x10000 - 50) {
            throw new RuntimeException("Invalid row number (" + totalRecords
                    + ") outside allowable range (0.." + (0x10000 - 50) + ")");
        }
    }


    private void createReportConditionSection(XLSSheetContext sheetContext) {
        HSSFWorkbook workbook = sheetContext.getWorkbook();
        HSSFSheet sheet = sheetContext.getRealSheet();
        int lastSheetRecordIndex = sheetContext.getAndIncrementLastRowIndex();
        if (!getConditions().isEmpty()) {
            HSSFRow reportConditionsTitleRow = sheet.createRow(lastSheetRecordIndex);
            sheet.addMergedRegion(new CellRangeAddress(lastSheetRecordIndex, lastSheetRecordIndex, 0, REPORT_TITLE_COL_SIZE - 1));
            HSSFCell reportConditionsTitleRowCell = reportConditionsTitleRow.createCell(0);
            reportConditionsTitleRowCell.setCellStyle(createPOIHeaderRowStyle(sheetContext, IndexedColors.LIGHT_ORANGE.getIndex(), false, true, true, true, 12));
            reportConditionsTitleRowCell.setCellValue(XLSUtils.getProperty(XLSConstants.DEFAULT_REPORT_CONDITIONS_TABLE_TITLE));
            HSSFRow conditionRow = null;
            int conditionColCount = 0;
            CellStyle reportConditionStyle = workbook.createCellStyle();
            for (XLSReportField condition : getConditions()) {
                if (conditionColCount == 0 || conditionColCount == REPORT_TITLE_COL_SIZE) {
                    conditionRow = sheet.createRow(sheetContext.getAndIncrementLastRowIndex());
                    conditionColCount = 0;
                }
                sheet.addMergedRegion(new CellRangeAddress(sheetContext.getLastSheetRecordIndex() - 1, sheetContext.getLastSheetRecordIndex() - 1,
                        conditionColCount, conditionColCount + 2));
                String value = condition.getValue();
                List values = condition.getValues();
                if (StringUtils.isEmpty(value)) {
                    value = "";
                }
                if (values == null) {
                    values = new ArrayList();
                    values.add(value);
                }

                if (values.size() == 1) {
                    HSSFCell label = conditionRow.createCell(conditionColCount);
                    String conditionCellValue = condition.getName() + " : " + values.get(0);
                    label.setCellValue(conditionCellValue);
                    int conditionRowHeight = computeRowHeight(conditionCellValue, 3);
                    if (conditionRowHeight > conditionRow.getHeight()) {
                        conditionRow.setHeight((short) conditionRowHeight);
                    }
                    label.setCellStyle(createPOIDataRowStyle(sheetContext, false, true, true, true, true, IndexedColors.LIGHT_YELLOW.getIndex(), reportConditionStyle));
                    for (int emptyCellIndex = conditionColCount + 1; emptyCellIndex <= conditionColCount + 2; emptyCellIndex++) {
                        HSSFCell emptyCell = conditionRow.createCell(emptyCellIndex, Cell.CELL_TYPE_STRING);
                        emptyCell.setCellValue("");
                        emptyCell.setCellStyle(createPOIDataRowStyle(sheetContext, false, true, true, true, true, IndexedColors.LIGHT_YELLOW.getIndex(), reportConditionStyle));
                    }
                } else {
                    //TODO
                  /*  for (Object conditionValue : values) {
                        HSSFCell label = conditionRow.createCell(conditionColCount);
                        label.setCellValue(condition.getName() + " :");
                        label.setCellStyle(createPOIDataRowStyle(sheetContext, false, true, true, true, true, IndexedColors.LIGHT_YELLOW.getIndex(), reportConditionStyle));

                    }*/
                }
                conditionColCount += 3;
            }

            if (conditionColCount != 0) {
                for (int remainCols = conditionColCount; remainCols < REPORT_TITLE_COL_SIZE; remainCols++) {
                    HSSFCell emptyCol = conditionRow.createCell(remainCols);
                    emptyCol.setCellValue("");
                    emptyCol.setCellStyle(createPOIDataRowStyle(sheetContext, false, true, true, true, true, IndexedColors.LIGHT_YELLOW.getIndex(), reportConditionStyle));
                }
                if (conditionColCount < REPORT_TITLE_COL_SIZE) {
                    sheet.addMergedRegion(new CellRangeAddress(sheetContext.getLastSheetRecordIndex() - 1,
                            sheetContext.getLastSheetRecordIndex() - 1, conditionColCount, REPORT_TITLE_COL_SIZE - 1));
                }
            } else {
                //TODO
            }
        }
    }

    private void addBorderToCellRange(HSSFWorkbook workbook, HSSFSheet sheet, CellRangeAddress cellRangeAddress) {
        HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, cellRangeAddress, sheet, workbook);
        HSSFRegionUtil.setTopBorderColor(IndexedColors.BLUE_GREY.getIndex(), cellRangeAddress, sheet, workbook);
        HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, cellRangeAddress, sheet, workbook);
        HSSFRegionUtil.setLeftBorderColor(IndexedColors.BLUE_GREY.getIndex(), cellRangeAddress, sheet, workbook);
        HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, cellRangeAddress, sheet, workbook);
        HSSFRegionUtil.setRightBorderColor(IndexedColors.BLUE_GREY.getIndex(), cellRangeAddress, sheet, workbook);
        HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, cellRangeAddress, sheet, workbook);
        HSSFRegionUtil.setBottomBorderColor(IndexedColors.BLUE_GREY.getIndex(), cellRangeAddress, sheet, workbook);
    }

    private void createEmptyRow(XLSSheetContext sheetContext) {
        HSSFSheet realSheet = sheetContext.getRealSheet();
        realSheet.createRow(sheetContext.getAndIncrementLastRowIndex());
    }

    private void createReportInformationRows(XLSSheetContext sheetContext) {
        HSSFSheet sheet = sheetContext.getRealSheet();
        List<XLSReportField> reportInformationFields = sheetContext.getReportInformationFields();
        int cellIndex = 0;
        int rowIndex = 0;
        int nextColIndex = 0;
        CellStyle style = sheetContext.getWorkbook().createCellStyle();
        HSSFRow reportInformationRow = null;
        for (XLSReportField reportField : reportInformationFields) {
            if (reportField.getWidth() > REPORT_TITLE_COL_SIZE) {
                throw new RuntimeException("Report Information field [ " + reportField.getName() + " ] width must be smaller than " + REPORT_TITLE_COL_SIZE);
            }
            if (nextColIndex >= REPORT_TITLE_COL_SIZE || nextColIndex == 0) {
                rowIndex = sheetContext.getAndIncrementLastRowIndex();
                reportInformationRow = sheet.createRow(rowIndex);
                setRowStyle(sheetContext, reportInformationRow, REPORT_TITLE_COL_SIZE, true, style);
                nextColIndex = 0;
            }
            HSSFCell branchInfoLabelCell = reportInformationRow.getCell(nextColIndex);
            branchInfoLabelCell.setCellValue(reportField.getValue());
            if (reportField.getWidth() > 1) {
                sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, nextColIndex, nextColIndex + reportField.getWidth() - 1));
            }
            nextColIndex += reportField.getWidth();
        }
        sheet.createRow(sheetContext.getAndIncrementLastRowIndex());
    }

    protected void setRowStyle(XLSSheetContext sheetContext, HSSFRow reportStatusFirstRow, int cellSize, boolean topBottomHeaderBorder, CellStyle style) {
        for (int cellIndex = 0; cellIndex < cellSize; cellIndex++) {
            HSSFCell cell = reportStatusFirstRow.createCell(cellIndex, Cell.CELL_TYPE_STRING);
            cell.setCellStyle(createPOIDataRowStyle(sheetContext, false, topBottomHeaderBorder, !topBottomHeaderBorder, false, false, IndexedColors.LIGHT_GREEN.getIndex(), style));
            cell.setCellValue("");
        }
    }

    private void createReportTitleRow(XLSSheetContext sheetContext) {
        HSSFSheet sheet = sheetContext.getRealSheet();
        HSSFRow titleRow = sheet.createRow(sheetContext.getAndIncrementLastRowIndex());
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, REPORT_TITLE_COL_SIZE - 1));
        HSSFCell titleRowCell = titleRow.createCell(0);
        titleRowCell.setCellStyle(createPOIHeaderRowStyle(sheetContext, IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex(), true, true, true, true, 14));
        titleRowCell.setCellValue(sheetContext.getReportTitle());
    }

    public void setRowCustomizerForAllSections(int sheetNo, XLSRowCustomizer rowCustomizer) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.setRowCustomizer(rowCustomizer);
    }

    public void setColorsDescription(int sheetNo, List<XLSColorDescription> colorsDescription) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.setColorsDescription(colorsDescription);
    }

    public void addCondition(XLSReportField reportConditionField) {
        if (reportConditionField == null) {
            throw new IllegalArgumentException("report condition field cant not be null!");
        }
        getConditions().add(reportConditionField);
    }

    public void addReportInformationFields(int sheetNo, XLSReportField reportInfoField) {
        XLSSheetContext sheetContext = getSheetContext(sheetNo);
        sheetContext.getReportInformationFields().add(reportInfoField);
    }

}
