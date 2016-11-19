package csv;

import ir.dotin.utils.csv.CSVReader;
import org.apache.commons.io.IOUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.junit.Assert;
import org.junit.Test;

import java.io.InputStream;
import java.util.List;
import java.util.Map;

/**
 * Created by r.rastakfard on 9/27/2016.
 */
public class CSVReaderTest {

    private static final Log log = LogFactory.getLog(CSVReaderTest.class);

    @Test
    public void testCSVReader() throws Exception {
        InputStream resourceAsStream = getClass().getResourceAsStream("/csv/sample-read.csv");
        String csvText;
        try {
            csvText = IOUtils.toString(resourceAsStream, "UTF-8");
        } finally {
            IOUtils.closeQuietly(resourceAsStream);
        }
        CSVReader obj = new CSVReader(csvText, true);
        List<Map<String, String>> parsedLines = obj.parse();
        Assert.assertEquals(8, parsedLines.size());
        for (Map<String, String> record : parsedLines) {
            for (String fieldName : record.keySet()) {
                log.info(fieldName + " : " + record.get(fieldName));
            }
            System.out.println();
        }

    }
}
