package ir.dotin.utils.xls.domain.builder;

import ir.dotin.utils.xls.domain.*;
import ir.dotin.utils.xls.renderer.XLSDefaultRowCustomizer;
import org.apache.commons.lang.StringUtils;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by r.rastakfard on 7/12/2016.
 */
public class XLSCellStyleBuilder<B> {
    Map<String, XLSCellStyle> result;
    Map<String, Short> fontsColor;
    Map<String, String> fontsName;
    private short defaultRowBackgroundColor;
    private short defaultRowFontColor;
    private String defaultRowFont;
    private XLSSheetContext sheetContext;

    public XLSCellStyleBuilder(XLSSheetContext sheetContext) {
        this.sheetContext = sheetContext;
        result = new HashMap<String, XLSCellStyle>();
        fontsColor = new HashMap<String, Short>();
        fontsName = new HashMap<String, String>();
    }

    public XLSCellStyleBuilder addCellDesignWithFormat(String cellName, String format) {
        if (StringUtils.isEmpty(cellName)) {
            throw new IllegalArgumentException("Cell Name is Empty!");
        }
        if (StringUtils.isEmpty(format)) {
            throw new IllegalArgumentException("Cell format is Empty!");
        }
        XLSCellStyle cellDesign = result.get(cellName);
        if (cellDesign==null) {
            boolean even = (sheetContext.getProcessedEntityCount()) % 2 == 0;
            String colorKey = even ? XLSConstants.DEFAULT_EVEN_ROW_COLOR_KEY : XLSConstants.DEFAULT_ODD_ROW_COLOR_KEY;
            XLSColorDescription bgColorDescription = sheetContext.getColorDescription(colorKey);
            XLSColorDescription fontColorDescription = sheetContext.getColorDescription(XLSConstants.DEFAULT_FONT_COLOR_KEY);
            cellDesign = new XLSCellStyle(new XLSCellFont(fontColorDescription, XLSConstants.DEFAULT_FONT_NAME), bgColorDescription, format);
        }else{
            cellDesign.setFormat(format);
        }
        result.put(cellName, cellDesign);
        return this;
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
        if (!fontsColor.isEmpty()) {
            XLSCellStyle xlsCellStyle;
            for (String colName : fontsColor.keySet()) {
                if (result.containsKey(colName)){
                    xlsCellStyle = result.get(colName);
                }else{
                    xlsCellStyle = XLSDefaultRowCustomizer.getDefaultCellStyle(sheetContext);
                    result.put(colName,xlsCellStyle);
                }
                xlsCellStyle.setFontColor(fontsColor.get(colName));
            }
        }
        if (!fontsName.isEmpty()) {
            XLSCellStyle xlsCellStyle;
            for (String colName : fontsName.keySet()) {
                if (result.containsKey(colName)){
                    xlsCellStyle = result.get(colName);
                }else{
                    xlsCellStyle = XLSDefaultRowCustomizer.getDefaultCellStyle(sheetContext);
                    result.put(colName,xlsCellStyle);
                }
                xlsCellStyle.setFontName(fontsName.get(colName));
            }
        }
        return result;
    }

    public XLSCellStyleBuilder addCellDesignWithFontColor(String cellName, short fontColor) {
        if (StringUtils.isEmpty(cellName)) {
            throw new IllegalArgumentException("Cell Name is Empty!");
        }
        if (StringUtils.isEmpty(cellName)) {
            throw new IllegalArgumentException("Font color is Empty!");
        }
        fontsColor.put(cellName, fontColor);
        return this;
    }

    public short getDefaultRowBackgroundColor() {
        return defaultRowBackgroundColor;
    }

    public void setDefaultRowBackgroundColor(short defaultRowBackgroundColor) {
        this.defaultRowBackgroundColor = defaultRowBackgroundColor;
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
        fontsName.put(colName, fontName);
        return this;
    }

    public XLSCellStyle createCellDesignWithBackgroundColor(String colorKey) {
        XLSColorDescription colorDescription = sheetContext.getColorDescription(colorKey);
        XLSCellStyle defaultCellDesign = XLSDefaultRowCustomizer.getDefaultCellStyle(sheetContext);
        defaultCellDesign.setBackGroundColor(colorDescription);
        return defaultCellDesign;
    }

    public XLSCellStyleBuilder addCellDesignWithBasicInfo(String colName) {

        return this;
    }

    public XLSCellStyleBuilder addCellDesignWithBasicInfo(String colName, String basicInfoKey) {
        return this;
    }

    public XLSCellStyleBuilder addCellDesignWithBasicInfo(String colName, List<B> basicInfoCollection) {
        return this;
    }

}
