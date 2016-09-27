package ir.dotin.utils.xls.mapper;

import ir.dotin.utils.xls.domain.XLSColumnDefinition;
import ir.dotin.utils.xls.domain.XLSRecord;
import ir.dotin.utils.xls.domain.XLSSheetContext;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Created by r.rastakfard on 7/27/2016.
 */
public class XLSReflectiveRowToEntityMapper<E> implements XLSRowToEntityMapper<E> {

    private Class mainEntityClass;

    public XLSReflectiveRowToEntityMapper(Class mainEntityClass) {
        this.mainEntityClass = mainEntityClass;
    }

    public E map(XLSRecord record, XLSSheetContext sheetContext) throws Exception {
        Class cls = mainEntityClass;
        E entity = null;
        try {
            entity = (E) cls.newInstance();
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        Map<String, List<Map<String, String>>> recordData = record.getRecordData();
        List<XLSColumnDefinition> recordColumns = record.getRecordColumns();
        for (XLSColumnDefinition definition : recordColumns) {
            String fieldName = definition.getName();
            if (definition.isRealColumn()) {
                List<Map<String, String>> fieldDataListMap = recordData.get(fieldName);
                List<String> fieldDataList = convertToDataList(fieldName, fieldDataListMap);
                try {
                    Method method = cls.getDeclaredMethod("set"+fieldName, definition.getFieldType());
                    method.invoke(entity, fieldDataList);
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                } catch (InvocationTargetException e) {
                    e.printStackTrace();
                } catch (NoSuchMethodException e) {
                    e.printStackTrace();
                }
            } else {
                for (XLSColumnDefinition subDefinition : definition.getSubColumns()){

                }

            }
        }
        return null;
    }

    private List<String> convertToDataList(String fieldName, List<Map<String, String>> fieldDataListMap) {
        List<String> result = new ArrayList<String>();
        for (Map<String, String> value : fieldDataListMap) {
            result.add(value.get(fieldName));
        }
        return result;
    }
}
