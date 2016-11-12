package ir.dotin.utils.xls.renderer;

import ir.dotin.utils.xls.domain.XLSColumnDefinition;
import ir.dotin.utils.xls.domain.XLSListReportSection;
import ir.dotin.utils.xls.domain.XLSReportField;
import ir.dotin.utils.xls.domain.XLSSheetContext;
import ir.dotin.utils.xls.mapper.XLSEntityToRowMapper;
import ir.dotin.utils.xls.writer.XLSListWriter;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

/**
 * Created by r.rastakfard on 7/14/2016.
 */
public class XLSListSectionRenderer extends XLSListWriter implements XLSSectionRenderer<XLSListReportSection> {

    public void renderSection(XLSSheetContext sheetContext, XLSListReportSection section) {
        createHeaderPartForSection(section, sheetContext);
        createXLSRecords(section, sheetContext);
    }

    private void createHeaderPartForSection(XLSListReportSection section, XLSSheetContext sheetContext) {
        HSSFSheet realSheet = createSectionTitleRow(section, sheetContext);
        createSectionAdditionalFieldsRow(section, sheetContext, realSheet);
        createHeaderRow(sheetContext,section.getHeaderCols());
    }

    private void createSectionAdditionalFieldsRow(XLSListReportSection section, XLSSheetContext sheetContext, HSSFSheet realSheet) {
        List<XLSReportField> titleFields = section.getTitleFields();
        if (titleFields != null && !titleFields.isEmpty()) {
            int colCount = 0;
            HSSFRow fieldsRow = null;
            for (int cellIndex = 1; cellIndex <= titleFields.size(); cellIndex++) {
                if (colCount == 0 || colCount == sheetContext.getReportGenerator().REPORT_TITLE_COL_SIZE) {
                    fieldsRow = realSheet.createRow(sheetContext.getAndIncrementLastRowIndex());
                    colCount = 0;
                }
                HSSFCell cell = fieldsRow.createCell(colCount, Cell.CELL_TYPE_STRING);
                realSheet.addMergedRegion(new CellRangeAddress(sheetContext.getLastSheetRecordIndex() - 1, sheetContext.getLastSheetRecordIndex() - 1,
                        colCount, colCount + 2));
                cell.setCellValue(titleFields.get(cellIndex - 1).getName() + " : " + titleFields.get(cellIndex - 1).getValue());
                cell.setCellStyle(createPOIHeaderRowStyle(sheetContext, IndexedColors.AQUA.getIndex(), true, true, true, true,12));

                for (int emptyCellIndex = colCount + 1; emptyCellIndex <= colCount + 2; emptyCellIndex++) {
                    HSSFCell emptyCell = fieldsRow.createCell(emptyCellIndex, Cell.CELL_TYPE_STRING);
                    emptyCell.setCellValue("");
                    emptyCell.setCellStyle(createPOIHeaderRowStyle(sheetContext, IndexedColors.AQUA.getIndex(), true, true, true, true, 12));
                }
                colCount += 3;
            }
        }
    }

    private HSSFSheet createSectionTitleRow(XLSListReportSection section, XLSSheetContext sheetContext) {
        HSSFSheet realSheet = sheetContext.getRealSheet();
        sheetContext.getAndIncrementLastRowIndex();
        if (StringUtils.isNotEmpty(section.getTitle())) {
            HSSFRow sectionTitleRow = realSheet.createRow(sheetContext.getAndIncrementLastRowIndex());
            HSSFCell cell = sectionTitleRow.createCell(0, Cell.CELL_TYPE_STRING);
            cell.setCellValue(section.getTitle());
            cell.setCellStyle(createPOIHeaderRowStyle(sheetContext, IndexedColors.AQUA.getIndex(), true, true, true, true, 12));
            sectionTitleRow.createCell(1, Cell.CELL_TYPE_STRING).setCellStyle(createPOIHeaderRowStyle(sheetContext, IndexedColors.AQUA.getIndex(), true, true, true, true, 12));
            sectionTitleRow.createCell(2, Cell.CELL_TYPE_STRING).setCellStyle(createPOIHeaderRowStyle(sheetContext, IndexedColors.AQUA.getIndex(), true, true, true, true, 12));
            sectionTitleRow.createCell(3, Cell.CELL_TYPE_STRING).setCellStyle(createPOIHeaderRowStyle(sheetContext, IndexedColors.AQUA.getIndex(), true, true, true, true, 12));
            sheetContext.getRealSheet().addMergedRegion(new CellRangeAddress(sheetContext.getLastSheetRecordIndex() - 1, sheetContext.getLastSheetRecordIndex() - 1, 0, 3));
        }
        return realSheet;
    }

    private void createXLSRecords(XLSListReportSection section, XLSSheetContext sheetContext) {
        List<XLSColumnDefinition> columnsDefinition = section.getHeaderCols();
        XLSEntityToRowMapper entityToRowMapper = section.getEntityToRowMapper();
        addDummyRecords(sheetContext,entityToRowMapper);
        XLSRowCustomizer rowCustomizer = section.getRowCustomizer();
        sheetContext.setEmptyRecordsMessage(section.getEmptyRecordsMessage());
        parseRecords(sheetContext, columnsDefinition, entityToRowMapper,rowCustomizer);
    }

}
