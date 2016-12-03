package xls;

import ir.dotin.test.domain.AddressVO;
import ir.dotin.test.domain.CustomerVO;
import ir.dotin.utils.xls.domain.*;
import ir.dotin.utils.xls.domain.builder.XLSCellStyleBuilder;
import ir.dotin.utils.xls.domain.builder.XLSColumnDefinitionBuilder;
import ir.dotin.utils.xls.domain.builder.XLSRowCellsDataBuilder;
import ir.dotin.utils.xls.mapper.XLSEntityToRowMapper;
import ir.dotin.utils.xls.renderer.XLSRowCustomizer;
import ir.dotin.utils.xls.writer.ExcelReportGenerator;
import ir.dotin.utils.xls.writer.XLSListWriter;
import org.apache.commons.io.FileUtils;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

/**
 * Created by r.rastakfard on 7/21/2016.
 */
public class XLSListWriterTest {

    public static String dummyExcelFile1Path = System.getProperty("java.io.tmpdir") + "testListWriter" + System.currentTimeMillis() + ".xls";
    public static String dummyExcelFile2Path = System.getProperty("java.io.tmpdir") + "testExcelReportGenerator" + System.currentTimeMillis() + ".xls";

    @BeforeClass
    public static void deleteDummyFiles() throws Exception {
        FileUtils.forceDeleteOnExit(new File(dummyExcelFile1Path));
        FileUtils.forceDeleteOnExit(new File(dummyExcelFile2Path));
    }

    @Test
    public void testListWriter() throws Exception {

        final List<String> locations = createLocationBasiceInfo();
        final List<String> nationalCodes = createnationalCodesBasiceInfo();

        XLSListWriter xlsWriter = new XLSListWriter<CustomerVO>();
        xlsWriter.setBasicInfoSheetProtectionPassword("yourOwnPassword");
        xlsWriter.addBasicInfoCollection("nationalCodesInfo", "انتخاب کد ملی", "کد ملی وارد شده در لیست وجود ندارد", nationalCodes);
        xlsWriter.addBasicInfoCollection("locationBasicInfo", "انتخاب شهر", "شهر انتخابی در لیست وجود ندارد", locations);
        xlsWriter.setBasicInfoSheetProtectionPassword("test1");
        initSheet(xlsWriter, 0, "لیست پرسنل");
//        initSheet(xlsWriter,1,"لیست پرسنل2");
        FileOutputStream outputStream = xlsWriter.createDocument(dummyExcelFile1Path);

    }

    private List<String> createnationalCodesBasiceInfo() {
        List<String> codes = new ArrayList<String>();
        codes.add("0074528637");
        codes.add("1521452244");
        codes.add("5845455454");
        codes.add("6998653313");
        return codes;
    }

    private void initSheet(XLSListWriter xlsWriter, int sheetIndex, String sheetName) {
        xlsWriter.setSheetName(sheetIndex, sheetName);
//        xlsWriter.setSheetAsReadonly(sheetIndex, "yourOwnPassword");
        List<XLSColumnDefinition> columnsDefinition = new XLSColumnDefinitionBuilder()
                .addColumnDefinition("customerNumber", "شماره مشتری")
                .addColumnDefinitionWithBasicInfo("nationalCode", "کد ملی", "nationalCodesInfo")
//                .addColumnDefinitionWithBasicInfo("birthLocation", "محل تولد", "locationBasicInfo")
                .addColumnDefinitionWithBasicInfo("birthLocation", "محل تولد", createLocationBasiceInfo())
//                .addColumnDefinitionWithBasicInfo("birthLocation", "محل تولد", "locationBasicInfo4",locations)
                .build();

        xlsWriter.setColumnsDefinition(sheetIndex, columnsDefinition);
        xlsWriter.addColorsDescription(sheetIndex, "green", 0, 255, 0, "توصیف تستی 1");
        List<CustomerVO> customers = generateDummyCustomers();
        xlsWriter.setRecords(sheetIndex, customers);
        xlsWriter.setRightToLeft(sheetIndex, true);
        xlsWriter.setEntityToRowMapper(sheetIndex, new XLSEntityToRowMapper<CustomerVO>() {
            public XLSRowCellsData map(CustomerVO entity, XLSSheetContext sheetContext) {
                return new XLSRowCellsDataBuilder()
                        .addCellSimpleData("customerNumber", String.valueOf(entity.getCustomerNumber()))
                        .addCellSimpleData("nationalCode", String.valueOf(entity.getNationalCode()))
                        .addCellSimpleData("birthLocation", entity.getBirthLocation())
                        .build();
            }

            public Map<String, XLSCellStyle> getRecordDesign(CustomerVO entity, XLSSheetContext sheetContext) {
                return new XLSCellStyleBuilder(sheetContext)
//                        .addCellDesignWithBasicInfo("birthLocation")
//                        .showCellDesignWithBasicInfo("birthLocation",false)
//                        .addCellDesignWithBasicInfo("birthLocation","locationBasicInfo")
//                        .addCellDesignWithBasicInfo("birthLocation",locations)
                        .build();
            }
        });
    }

    private List<String> createLocationBasiceInfo() {
        List<String> countryNames = new ArrayList<String>();
        countryNames.add("تهران");
        countryNames.add("تبریز");
        countryNames.add("همدان");
        countryNames.add("اصفهان");
        countryNames.add("ایلام");
        countryNames.add("قزوین");
        countryNames.add("ملایر");
        countryNames.add("رشت");
        countryNames.add("ساری");
        return countryNames;
    }

    @Test
    public void testExcelReportGenerator() throws Exception {
        List<CustomerVO> customers = generateDummyCustomers();

        // Initialize Generator
        ExcelReportGenerator generator = new ExcelReportGenerator();
        generator.setReportTitle(0, "گزارش لیست پرسنل");
        generator.setSheetName(0, "لیست پرسنل");

        generator.addBasicInfoCollection("nationalCodesInfo", "انتخاب کد ملی", "کد ملی وارد شده در لیست وجود ندارد", createnationalCodesBasiceInfo());
        generator.addColorsDescription(0, "red", 255, 0, 0, "توصیف 1");
        generator.addColorsDescription(0, "blue", 0, 0, 255, "توصیف 2");
        generator.addColorsDescription(0, "green", 0, 255, 0, "توصیف 3");
        generator.addColorsDescription(0, "violet", 255, 0, 255, "توصیف 4");

        generator.addCondition(new XLSReportField("نام", "سعید"));
        generator.addCondition(new XLSReportField("نام خانوادگی", "رستاک"));
        generator.addCondition(new XLSReportField("شماره پرسنلی", "123456"));
        generator.addCondition(new XLSReportField("نام پدر", "محمد"));
        generator.addCondition(new XLSReportField("محل تولد", "تهران"));

        XLSListReportSection xlsReportSection = new XLSListReportSection<CustomerVO>();

        xlsReportSection.setRowCustomizer(new XLSRowCustomizer() {
            public Map<String, XLSCellStyle> createRecordStyle(Map<String, List<Map<String, String>>> row, XLSSheetContext sheetContext) {
                return null;
            }
        });

        List<XLSColumnDefinition> addressSubColumns = new XLSColumnDefinitionBuilder()
                .addColumnDefinition("type", "نوع")
                .addColumnDefinition("address", "آدرس", 7)
                .build();
        List<XLSColumnDefinition> cols = new XLSColumnDefinitionBuilder()
                .addColumnDefinition("customerNumber", "شماره مشتری")
                .addColumnDefinitionWithBasicInfo("nationalCode", "کد ملی", "nationalCodesInfo")
                .addColumnDefinition("addresses", "آدرس", addressSubColumns)
                .build();
        xlsReportSection.setHeaderCols(cols);
        xlsReportSection.setRecords(customers);
        xlsReportSection.setEntityToRowMapper(new XLSEntityToRowMapper<CustomerVO>() {
            public XLSRowCellsData map(CustomerVO entity, XLSSheetContext sheetContext) {
                List<Map<String, String>> addresses = new ArrayList<Map<String, String>>();
                for (AddressVO address : entity.getAddresses()) {
                    Map<String, String> addressRecord = new HashMap<String, String>();
                    addressRecord.put("type", address.getType());
                    addressRecord.put("address", address.getAddress());
                    addresses.add(addressRecord);
                }
                XLSRowCellsData rowDefinition = new XLSRowCellsDataBuilder()
                        .addCellSimpleData("customerNumber", String.valueOf(entity.getCustomerNumber()))
                        .addCellSimpleData("nationalCode", entity.getNationalCode())
                        .addEntityListCellsData("addresses", addresses)
                        .build();
                return rowDefinition;
            }

            public Map<String, XLSCellStyle> getRecordDesign(CustomerVO entity, XLSSheetContext sheetContext) {
                return null;
            }
        });

        xlsReportSection.addTitleField(new XLSReportField("نام شعبه", "ستاد"));
        xlsReportSection.addTitleField(new XLSReportField("کد شعبه", "1"));

        generator.addReportSection(0, xlsReportSection);
        generator.createDocument(dummyExcelFile2Path);

    }

    private List<CustomerVO> generateDummyCustomers() {
        List<CustomerVO> result = new ArrayList<CustomerVO>();
        CustomerVO customer1 = new CustomerVO(1234l, "0074528637");
        AddressVO addressVO1 = new AddressVO("آدرس منزل", "تهران - خیابان انقلاب - خیابان اردیبهشت - پلاک 18");
        AddressVO addressVO11 = new AddressVO("محل کار", "تهران - خیابان انقلاب - خیابان اردیبهشت - پلاک 19");
        customer1.setAddresses(Arrays.asList(addressVO1, addressVO11));
        customer1.setBirthLocation("تهران");
        CustomerVO customer2 = new CustomerVO(5678l, "0065498714");
        AddressVO addressVO2 = new AddressVO("آدرس منزل", "تهران - خیابان انقلاب - خیابان فروردین - پلاک 12");
        customer2.setAddresses(Arrays.asList(addressVO2));
        customer2.setBirthLocation("تبریز");
        CustomerVO customer3 = new CustomerVO(25874l, "3300216547");
        AddressVO addressVO3 = new AddressVO("محل کار", "تهران - خیابان ونک - خیابان گاندی - پلاک 22");
        customer3.setAddresses(Arrays.asList(addressVO3));
        customer3.setBirthLocation("همدان");
        result.add(customer1);
        result.add(customer2);
        result.add(customer3);
        return result;
    }
}
