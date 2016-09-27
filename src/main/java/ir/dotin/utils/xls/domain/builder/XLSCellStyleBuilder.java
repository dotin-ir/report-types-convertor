package ir.dotin.utils.xls.domain.builder;

import ir.dotin.utils.xls.domain.XLSCellStyle;
import ir.dotin.utils.xls.domain.XLSColorDescription;
import ir.dotin.utils.xls.domain.XLSSheetContext;
import ir.dotin.utils.xls.renderer.XLSDefaultRowCustomizer;
import org.apache.commons.lang.StringUtils;

import java.util.HashMap;
import java.util.Map;

/**
 * Created by r.rastakfard on 7/12/2016.
 */
public class XLSCellStyleBuilder {
    Map<String, XLSCellStyle> result;
    private short defaultRowBackgroundColor;
    private short defaultRowFontColor;
    private String defaultRowFont;
    Map<String, String> fonts;
    private XLSSheetContext sheetContext;

    public XLSCellStyleBuilder(XLSSheetContext sheetContext) {
        this.sheetContext = sheetContext;
        result = new HashMap<String, XLSCellStyle>();
        fonts = new HashMap<String, String>();
    }

    public XLSCellStyleBuilder addCellDesignWithBackgroundColor(String cellName, XLSCellStyle cellDesign) {
        if (cellDesign == null) {
            throw new IllegalArgumentException("Cell Design is null!");
        }
        if (StringUtils.isEmpty(cellName)) {
            throw new IllegalArgumentException("Cell Name is Empty!");
        }
        result.put(cellName, cellDesign);
        return this;
    }

    public Map<String, XLSCellStyle> build() {
        return result;
    }

    public XLSCellStyleBuilder addCellDesignWithFontColor(String cellName, short fontColor) {
        return null;
    }

    public void setDefaultRowBackgroundColor(short defaultRowBackgroundColor) {
        this.defaultRowBackgroundColor = defaultRowBackgroundColor;
    }

    public short getDefaultRowBackgroundColor() {
        return defaultRowBackgroundColor;
    }

    public XLSCellStyleBuilder setDefaultRowFontColor(short defaultRowFontColor) {
        this.defaultRowFontColor = defaultRowFontColor;
        return this;
    }

    public XLSCellStyleBuilder setDefaultRowFont(String defaultRowFont) {
        if (StringUtils.isEmpty(defaultRowFont)) {
            throw new IllegalArgumentException("defaultRowFont is empty!");
        }
        this.defaultRowFont = defaultRowFont;
        return this;
    }

    public XLSCellStyleBuilder setCellDesignFont(String colName, String fontName) {
        if (StringUtils.isEmpty(colName)) {
            throw new IllegalArgumentException("Cell name is empty!");
        }
        if (StringUtils.isEmpty(fontName)) {
            throw new IllegalArgumentException("Cell name is empty!");
        }
        fonts.put(colName, fontName);
        return this;
    }

    public XLSCellStyle createCellDesignWithBackgroundColor(String colorKey) {
        XLSColorDescription colorDescription = sheetContext.getColorDescription(colorKey);
        XLSCellStyle defaultCellDesign = XLSDefaultRowCustomizer.getDefaultCellStyle(sheetContext);
        defaultCellDesign.setBackGroundColor(colorDescription);
        return defaultCellDesign;
    }
}
