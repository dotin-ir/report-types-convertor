package ir.dotin.utils.xls.checker;

import ir.dotin.utils.xls.domain.XLSColumnDefinition;
import ir.dotin.utils.xls.domain.XLSDuplicatePolicy;
import ir.dotin.utils.xls.domain.XLSRecord;
import ir.dotin.utils.xls.domain.XLSSheetContext;

import java.util.List;
import java.util.Map;

/**
 * Created by r.rastakfard on 6/23/2016.
 */
public class XLSReportAllAsErrorUniqueChecker extends XLSUniqueChecker {

    public XLSReportAllAsErrorUniqueChecker() {
    }

    @Override
    public boolean checkUniqueRow(XLSRecord currentXlsRecord, XLSSheetContext sheetContext) {
        List<XLSColumnDefinition> recordColumns = currentXlsRecord.getRecordColumns();
        Map<String, List<Map<String, String>>> currentRecord = currentXlsRecord.getRecordData();
        boolean valid = true;
        if (isKeySet(sheetContext)) {
            boolean existInRow = isUniqueColumnExistInRow(currentRecord, recordColumns);
            if (existInRow) {
                valid = validateUniqueRow(currentXlsRecord, sheetContext);
            } else {
                valid = false;
            }
        }
        super.onFinish(valid, sheetContext);
        return valid;
    }

    private boolean isUniqueColumnExistInRow(Map<String, List<Map<String, String>>> currentRecord, List<XLSColumnDefinition> recordColumns) {
        for (XLSColumnDefinition definition : recordColumns) {
            if (definition.isUniqueColumn()) {
                if (!currentRecord.containsKey(definition.getName())) {
                    return false;
                }
            }
            for (XLSColumnDefinition subDefinition : definition.getSubColumns()) {
                if (subDefinition.isUniqueColumn()) {
                    List<Map<String, String>> subRecordDataList = currentRecord.get(definition.getName());
                    for (Map<String, String> data : subRecordDataList) {
                        if (!data.containsKey(subDefinition.getName())) {
                            return false;
                        }
                    }
                }
            }
        }
        return true;
    }

    @Override
    public XLSDuplicatePolicy getDuplicatePolicy() {
        return XLSDuplicatePolicy.REPORT_ALL_AS_ERROR;
    }

    private boolean validateUniqueRow(final XLSRecord currentRecord, XLSSheetContext context) {
        List<XLSRecord> validRecords = context.getRawRecords();
        List<XLSRecord> errorRecords = context.getErrorRecords();
        String hash = getHash(currentRecord, context);
        if (context.isExistInErrorRecord(hash)) {
            context.addToErrorList(hash, currentRecord);
            return false;
        }
        for (XLSRecord record : validRecords) {
            if (compareRecord(record, currentRecord, context) == 0) {
                context.addToErrorList(hash);
                return false;
            }
        }
        for (XLSRecord record : errorRecords) {
            if (compareRecord(record, currentRecord, context) == 0) {
                context.addToErrorList(hash);
                return false;
            }
        }

        return true;
    }

    public String getHash(XLSRecord record, XLSSheetContext sheetContext) {
        String hashCode = "";
        List<String> uniqueColumnKey = sheetContext.getUniqueColumnKey();
        for (String key : uniqueColumnKey) {
            String data = getData(record, key);
            if (data != null) {
                hashCode += String.valueOf(data.hashCode());
            } else {
                hashCode += key;
            }
        }
        return hashCode;
    }

    private String getData(XLSRecord record, String key) {
        Map<String, List<Map<String, String>>> recordData = record.getRecordData();
        List<XLSColumnDefinition> recordColumns = record.getRecordColumns();
        for (XLSColumnDefinition definition : recordColumns) {
            if (!definition.isHidden() && definition.isRealColumn() && definition.getName().equals(key)) {
                if (recordData.containsKey(key)) {
                    List<Map<String, String>> maps = recordData.get(key);
                    if (maps.size() > 1) {
                        throw new RuntimeException("can not compute hash for list column(you can not set list column as unique key)!");
                    }
                    return maps.get(0).get(key);
                }
            }
        }
        return null;
    }
}
