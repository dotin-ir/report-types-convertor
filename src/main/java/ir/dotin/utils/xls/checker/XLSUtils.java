package ir.dotin.utils.xls.checker;

import ir.dotin.utils.xls.domain.XLSColumnDefinition;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.filefilter.FileFilterUtils;
import org.apache.commons.io.filefilter.IOFileFilter;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.MessageFormat;
import java.util.Collection;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

/**
 * Created by r.rastakfard on 7/24/2016.
 */
public class XLSUtils {

    private static String localizationFilePath = XLSUtils.class.getClassLoader().getResource("localization.properties").getPath().replace("%20", " ");
    private static long lastModified = 0;
    private static Properties properties = new Properties();
    private static Boolean init=false;

    static {
        initialize();
    }

    public static boolean initialize() {
        try {
            synchronized (init) {
                if (!init) {
                    File file = new File(localizationFilePath);
                    if (file.lastModified() != lastModified) {
                        properties.clear();
                        properties.load(new FileInputStream(localizationFilePath));
                        lastModified = file.lastModified();
                        init = true;
                    }
                }
            }
        } catch (IOException e) {
        }
        return init;
    }
    public static boolean hasSubColumns(XLSColumnDefinition xlsColumnDefinition) {
        return xlsColumnDefinition.getSubColumns() != null && !xlsColumnDefinition.getSubColumns().isEmpty();
    }

    public static int compareSimpleRecord(Map<String, String> rec1, Map<String, String> rec2) {
        if (rec1 == null) rec1 = new HashMap<String, String>();
        if (rec2 == null) rec2 = new HashMap<String, String>();
        if (rec1.size() == rec2.size()) {
            for (String key : rec1.keySet()) {
                String data1 = rec1.get(key);
                String data2 = rec2.get(key);
                if (data1 == null && data2 != null) {
                    return -1;
                }
                if (data1 != null && data2 == null) {
                    return 1;
                }
                if (data1 == null && data2 == null) {
                    continue;
                }
                return data1.compareTo(data2);
            }
        } else {
            return new Integer(rec1.size()).compareTo(rec2.size());
        }
        return 0;
    }

    public static boolean isValidNumber(String numberStr) {
        boolean result = true;
        try {
            BigDecimal.valueOf(Double.valueOf(numberStr));
        } catch (Exception ex) {
            result = false;
        }
        return result;
    }


    public static void copyRow(Sheet worksheet, int rowNum, Row sourceRow) {

        //Save the text of any formula before they are altered by row shifting
        String[] formulasArray = new String[sourceRow.getLastCellNum()];
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            if (sourceRow.getCell(i) != null && sourceRow.getCell(i).getCellType() == Cell.CELL_TYPE_FORMULA)
                formulasArray[i] = sourceRow.getCell(i).getCellFormula();
        }

        worksheet.shiftRows(rowNum, worksheet.getLastRowNum(), 1);
        Row newRow = worksheet.getRow(rowNum + 1); //Now sourceRow is the empty line, so let's rename it


        // Loop through source columns to add to new row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            Cell oldCell = sourceRow.getCell(i);
            Cell newCell;

            // If the old cell is null jump to next cell
            if (oldCell == null) {
                continue;
            } else {
                newCell = newRow.createCell(i);
            }

            // Copy style from old cell and apply to new cell
            CellStyle newCellStyle = worksheet.getWorkbook().createCellStyle();
//            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
            newCell.setCellStyle(newCellStyle);

            // If there is a cell comment, copy
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            // Set the cell data type
            newCell.setCellType(oldCell.getCellType());

            // Set the cell data value
            switch (oldCell.getCellType()) {
                case Cell.CELL_TYPE_BLANK:
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_ERROR:
                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    newCell.setCellFormula(formulasArray[i]);
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    newCell.setCellValue(oldCell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    newCell.setCellValue(oldCell.getRichStringCellValue());
                    break;
                default:
                    break;
            }
        }

        // If there are any merged regions in the source row, copy to new row
        for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
            CellRangeAddress cellRangeAddress = worksheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
                        (newRow.getRowNum() +
                                (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow()
                                )),
                        cellRangeAddress.getFirstColumn(),
                        cellRangeAddress.getLastColumn());
                worksheet.addMergedRegion(newCellRangeAddress);
            }
        }
    }


    public static void printPersianFontsInJavaFormat() {
        File file = new File("C:\\Windows\\Fonts");
        IOFileFilter fileFilter = FileFilterUtils.prefixFileFilter("B ");
        Collection<File> collection = FileUtils.listFiles(file, fileFilter, null);
        for (File fontFile : collection) {
            String fontFileName = fontFile.getName().substring(0, fontFile.getName().lastIndexOf("_"));
            String fieldName = fontFileName.replaceAll(" ", "_").toUpperCase();
            System.out.println("String " + fieldName + " = \"" + fontFileName + "\";");
        }
    }

    public static boolean isEmpty(String str) {
        if (str != null) {
            str = str.trim();
        }
        return StringUtils.isEmpty(str);
    }

    public static boolean isNotEmpty(String str) {
        return !isEmpty(str);
    }

    public static void main(String[] args) {
        XLSUtils.printPersianFontsInJavaFormat();
    }

    public static String refineCharacters(String str) {
        if (str != null) {
            str = str
                    .replaceAll("ی", "ي")
                    .replaceAll("ئ", "ي")
                    .replaceAll("ى", "ي")
                    .replaceAll("ء", "ي")
                    .replaceAll("أ", "ا")
                    .replaceAll("إ", "ا")
                    .replaceAll("ٱ", "ا")
                    .replaceAll("ك", "ک")
                    .replaceAll("ڪ", "ک")
                    .replaceAll("ؤ", "و")
                    .replaceAll("ة", "ه")
                    .replaceAll("\"", "")
                    .replaceAll("'", "")
                    .replaceAll("،", "")
                    .replaceAll("ٰ", "")
                    .replaceAll("؛", "")
                    .replaceAll("ٔ", "")
                    .replaceAll("ً", "")
                    .replaceAll("ٌ", "")
                    .replaceAll("ٍ", "")
                    .replaceAll("ّ", "")
                    .replaceAll("َ", "")
                    .replaceAll("ُ", "")
                    .replaceAll("ِ", "")
                    .replaceAll("ْ", "")
            ;
            str = str.trim();
        }
        return str;
    }

    public static String getProperty(String key) {
        return (String) properties.get(key);
    }

    public static String getPropertyWithPattern(String key, String... params) {
        String s = (String) properties.get(key);
        return MessageFormat.format(s, params);
    }

    public static String getProperty(String key, String subKey) {
        String secondaryKey = MessageFormat.format(key, subKey);
        return getProperty(secondaryKey);
    }
}
