package ir.dotin.utils.xls.domain;

import org.apache.commons.lang.builder.EqualsBuilder;
import org.apache.commons.lang.builder.HashCodeBuilder;

import java.io.Serializable;

/**
 * Created by r.rastakfard on 8/9/2016.
 */
public class XLSCellFont implements Serializable {
    private XLSColorDescription fontColor;
    private String fontName;
    private int size = XLSConstants.DEFAULT_CELL_FONT_SIZE;
//    private int charset

    public XLSCellFont(XLSColorDescription fontColor, String fontName) {
        this.fontColor = fontColor;
        this.fontName = fontName;
    }

    public XLSColorDescription getFontColor() {
        return fontColor;
    }

    public void setFontColor(XLSColorDescription fontColor) {
        this.fontColor = fontColor;
    }

    public String getFontName() {
        return fontName;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;

        if (!(o instanceof XLSCellFont)) return false;

        XLSCellFont that = (XLSCellFont) o;

        return new EqualsBuilder()
                .append(fontColor, that.fontColor)
                .append(fontName, that.fontName)
                .isEquals();
    }

    @Override
    public int hashCode() {
        return new HashCodeBuilder(17, 37)
                .append(fontColor)
                .append(fontName)
                .append(size)
                .toHashCode();
    }

    public int getSize() {
        return size;
    }

    public void setSize(int size) {
        this.size = size;
    }
}
