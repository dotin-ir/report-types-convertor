package ir.dotin.utils.xls.domain;

import org.apache.commons.lang.builder.EqualsBuilder;
import org.apache.commons.lang.builder.HashCodeBuilder;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;

import java.io.Serializable;

/**
 * Created by r.rastakfard on 7/12/2016.
 */
public class XLSCellStyle implements Serializable {
    private XLSCellFont font;
    private XLSColorDescription backGroundColor;
    private String format = XLSConstants.DEFAULT_CELL_FORMAT;
    private Boolean hasTopBorder = Boolean.TRUE;
    private Boolean hasRightBorder = Boolean.TRUE;
    private Boolean hasLeftBorder = Boolean.TRUE;
    private Boolean hasBottomBorder = Boolean.TRUE;
    private Boolean defaultDesign = Boolean.FALSE;
    private HSSFCellStyle realCellStyle;


    public XLSCellStyle(XLSCellFont font, XLSColorDescription backGroundColor, String format) {
        this.font = font;
        this.backGroundColor = backGroundColor;
        this.format = format;
    }

    public XLSColorDescription getBackGroundColor() {
        return backGroundColor;
    }

    public void setBackGroundColor(XLSColorDescription backGroundColor) {
        this.backGroundColor = backGroundColor;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    public Boolean getHasTopBorder() {
        return hasTopBorder;
    }

    public void setHasTopBorder(Boolean hasTopBorder) {
        this.hasTopBorder = hasTopBorder;
    }

    public Boolean getHasRightBorder() {
        return hasRightBorder;
    }

    public void setHasRightBorder(Boolean hasRightBorder) {
        this.hasRightBorder = hasRightBorder;
    }

    public Boolean getHasLeftBorder() {
        return hasLeftBorder;
    }

    public void setHasLeftBorder(Boolean hasLeftBorder) {
        this.hasLeftBorder = hasLeftBorder;
    }

    public Boolean getHasBottomBorder() {
        return hasBottomBorder;
    }

    public void setHasBottomBorder(Boolean hasBottomBorder) {
        this.hasBottomBorder = hasBottomBorder;
    }

    public Boolean getDefaultDesign() {
        return defaultDesign;
    }

    public void setDefaultDesign(Boolean defaultDesign) {
        this.defaultDesign = defaultDesign;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;

        if (!(o instanceof XLSCellStyle)) return false;

        XLSCellStyle that = (XLSCellStyle) o;

        return new EqualsBuilder()
                .append(font, that.font)
                .append(backGroundColor, that.backGroundColor)
                .append(format, that.format)
                .isEquals();
    }

    @Override
    public int hashCode() {
        return new HashCodeBuilder(17, 37)
                .append(font)
                .append(backGroundColor)
                .append(format)
                .toHashCode();
    }

    public HSSFCellStyle getRealCellStyle() {
        return realCellStyle;
    }

    public void setRealCellStyle(HSSFCellStyle realCellStyle) {
        this.realCellStyle = realCellStyle;
    }

    public XLSCellFont getFont() {
        return font;
    }

    public void setFont(XLSCellFont font) {
        this.font = font;
    }
}
