import ir.dotin.utils.xls.checker.XLSUtils;
import ir.dotin.utils.xls.domain.XLSConstants;
import org.junit.Assert;
import org.junit.Test;

/**
 * Created by r.rastakfard on 8/10/2016.
 */
public class LocalizationTest {
    @Test
    public void testLocalization() throws Exception {
        Assert.assertEquals("Error in localizaton", XLSUtils.getProperty(XLSConstants.DEFAULT_FONT_COLOR_DESCRIPTION), "رنگ پیش فرض فونت");
    }
}
