package ir.dotin.utils.xls.domain;

import ir.dotin.utils.xls.writer.XLSBaseWriter;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.io.Serializable;
import java.util.List;

/**
 * Created by r.rastakfard on 11/22/2016.
 */
public class XLSBasicInfo<B> implements Serializable {
    private final String basicInfoKey;
    private final String errorBoxTitle;
    private final String errorBoxMessage;
    private final List<B> collection;
    private DataValidation validation;
    private int cellIndex;
    private DVConstraint constraint;

    public XLSBasicInfo(String basicInfoKey, String errorBoxTitle, String errorBoxMessage, List<B> collection) {
        this.basicInfoKey = basicInfoKey;
        this.errorBoxTitle = errorBoxTitle;
        this.errorBoxMessage = errorBoxMessage;
        this.collection = collection;
    }

    public List<B> getCollection() {
        return collection;
    }

    public String getBasicInfoKey() {
        return basicInfoKey;
    }

    public String getErrorBoxTitle() {
        return errorBoxTitle;
    }

    public String getErrorBoxMessage() {
        return errorBoxMessage;
    }

    public DataValidation createValidation(CellRangeAddressList addressList, XLSBaseWriter writer) {
        if (!writer.hasSheet(XLSConstants.BASIC_INFO_SHEET_NAME)) {
            createBasicInfoSheet(writer);
        }
        if (constraint == null) {
            constraint = createConstraint(writer);
            writer.createBasicInfoColumn(writer.getWorkBook().getSheet(XLSConstants.BASIC_INFO_SHEET_NAME), getCellIndex(), getBasicInfoKey(), this);
        }
        validation = createDataValidation(addressList, constraint);
        return validation;
    }

    private DataValidation createDataValidation(CellRangeAddressList addressList, DVConstraint constraint) {
        DataValidation validation = new HSSFDataValidation(addressList, constraint);
        validation.createErrorBox(getErrorBoxTitle(), getErrorBoxMessage());
        return validation;
    }

    private DVConstraint createConstraint(XLSBaseWriter writer) {
        Name namedCell = writer.getWorkBook().createName();
        namedCell.setNameName(getBasicInfoKey());
        namedCell.setRefersToFormula(XLSConstants.BASIC_INFO_SHEET_NAME + "!$" + CellReference.convertNumToColString(getCellIndex())
                + "$2:$" + CellReference.convertNumToColString(getCellIndex()) + "$" + (getCollection().size() + 1));
        return DVConstraint.createFormulaListConstraint(getBasicInfoKey());
    }

    private void createBasicInfoSheet(XLSBaseWriter writer) {
        HSSFSheet sheet = writer.getWorkBook().createSheet(XLSConstants.BASIC_INFO_SHEET_NAME);
        sheet.setRightToLeft(true);
        sheet.protectSheet(writer.getBasicInfoSheetProtectionPassword());
        writer.getWorkBook().setSheetHidden(writer.getWorkBook().getSheetIndex(sheet), true);
    }

    public void setValidation(DataValidation validation) {
        this.validation = validation;
    }

    public void setCellIndex(int cellIndex) {
        this.cellIndex = cellIndex;
    }

    public int getCellIndex() {
        return cellIndex;
    }
}
