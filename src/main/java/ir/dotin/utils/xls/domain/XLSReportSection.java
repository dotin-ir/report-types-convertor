package ir.dotin.utils.xls.domain;

import ir.dotin.utils.xls.checker.XLSUtils;
import ir.dotin.utils.xls.renderer.XLSSectionRenderer;

import java.io.Serializable;

/**
 * Created by r.rastakfard on 7/14/2016.
 */
public abstract class XLSReportSection  implements Serializable {

    private String emptyRecordsMessage = XLSConstants.DEFAULT_EMPTY_SECTION_RECORDS_MESSAGE;

    public abstract XLSSectionRenderer getRenderer();

    public abstract Integer getRecordsCount();

    public void setEmptyRecordsMessage(String emptyRecordsMessage) {
        this.emptyRecordsMessage = emptyRecordsMessage;
    }

    public String getEmptyRecordsMessage() {
        return emptyRecordsMessage;
    }
}
