package ir.dotin.utils.csv;

import org.apache.commons.lang.StringUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by r.rastakfard on 2/10/2016.
 */
public class CSVReader {

    private static final Log log = LogFactory.getLog(CSVReader.class);
    private String cvsString;
    private String[] fields;
    private boolean hasFieldName = false;

    public CSVReader(String cvsText, boolean hasFieldName) {
        this.cvsString = cvsText;
        this.hasFieldName = hasFieldName;
    }

    public List<Map<String, String>> parse() {

        if (StringUtils.isEmpty(cvsString)) return new ArrayList<Map<String, String>>();

        String cvsSplitBy = ",";
        log.info("Start Parsing CVS file...");
        Map<String, String> records = new HashMap<String, String>();
        List<Map<String, String>> result = new ArrayList<Map<String, String>>();
        String[] splitLines = cvsString.split("\n");
        int firstLineNumber = 0;
        if (hasFieldName) {
            for (int j = 0; j < splitLines.length; j++) {
                if (!splitLines[j].contains(cvsSplitBy)) {
                    continue;
                } else {
                    firstLineNumber = j;
                    fields = removeUnknownChars(splitLines[j].split(cvsSplitBy));
                    break;
                }
            }
        }
        int rowNumber = 1;
        for (int i = firstLineNumber; i < splitLines.length; i++) {
            records = new HashMap<String, String>();
            if (hasFieldName && i == firstLineNumber) {
                log.info(splitLines[firstLineNumber]);
                continue;
            }
            String line = splitLines[i].trim();
            // use comma as separator
            if (!line.contains(cvsSplitBy)) {
                continue;
            }
            String[] values = line.split(cvsSplitBy);
            String resultRowStr = "";
            for (int j = 0; j < values.length; j++) {
                if (hasFieldName) {
                    records.put(fields[j].trim(), values[j].trim());
                    resultRowStr += values[j].trim() + ",";
                } else {
                    records.put(String.valueOf(j), values[j].trim());
                }

            }
            log.info("Record " + (rowNumber) + " [ " + resultRowStr + " ]");
            rowNumber++;
            result.add(records);
        }
        return result;

    }

    private String[] removeUnknownChars(String[] fields) {
        for (int k = 0; k < fields.length; k++) {
            String field = fields[k];
            char[] c2 = new char[field.length()];
            int l = 0;
            int e = 0;
            for (char c : field.toCharArray()) {
                char[] c1 = new char[1];
                c1[0] = c;
                if ("qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890-_.".contains(new String(c1))) {
                    c2[l] = c;
                    l++;
                } else {
                    e++;
                }
            }
            char[] f = new char[c2.length - e];
            for (int z = 0; z < c2.length - e; z++) {
                f[z] = c2[z];

            }
            fields[k] = new String(f);
        }
        return fields;
    }

    public boolean hasField(String fieldName) {
        if (StringUtils.isEmpty(fieldName)) {
            return false;
        }
        if (fields.length > 0) {
            for (String field : fields) {
                if (field.trim().toString().equals(fieldName.trim())) {
                    return true;
                }
            }
        }
        return false;
    }

}