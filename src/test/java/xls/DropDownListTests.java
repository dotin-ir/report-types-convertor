package xls;

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;

/**
 * Created by r.rastakfard on 11/16/2016.
 */
public class DropDownListTests {

    @Test
    public void testDropDownListWithHiddenSheet() throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet realSheet = workbook.createSheet("Sheet xls");
        HSSFSheet hidden = workbook.createSheet("hiddenSheet");
        String[] countryName = {"تهران", "تبریز", "همدان", "اصفهان"};
        for (int i = 0, length = countryName.length; i < length; i++) {
            String name = countryName[i];
            HSSFRow row = hidden.createRow(i);
            HSSFCell cell = row.createCell(0);
            cell.setCellValue(name);
        }
        Name namedCell = workbook.createName();
        namedCell.setNameName("hidden");
        namedCell.setRefersToFormula("hiddenSheet!$A$1:$A$" + countryName.length);
        DVConstraint constraint = DVConstraint.createFormulaListConstraint("hidden");


        CellRangeAddressList addressList = new CellRangeAddressList(0, 99, 0, 0);
        HSSFDataValidation validation = new HSSFDataValidation(addressList, constraint);
        validation.createErrorBox("انتخاب شهر", "شهر انتخابی در لیست وجود ندارد");
        realSheet.addValidationData(validation);

/*        CellRangeAddressList addressList2 = new CellRangeAddressList(1, 1, 0, 0);
        HSSFDataValidation validation2 = new HSSFDataValidation(addressList2, constraint);
        validation2.createErrorBox("انتخاب شهر", "شهر انتخابی در لیست وجود ندارد");
        realSheet.addValidationData(validation2);*/

        workbook.setSheetHidden(1, true);

//        workbook.
        realSheet.createRow(0).createCell(0).setCellValue("تهران");
        realSheet.createRow(1).createCell(0).setCellValue("تبریز");
        FileOutputStream stream = new FileOutputStream("d:\\range.xls");
        workbook.write(stream);
        stream.close();

      /*  Biff8EncryptionKey.setCurrentUserPassword("pass");
        NPOIFSFileSystem fs = new NPOIFSFileSystem(new File("file.xls"), true);
        HSSFWorkbook hwb = new HSSFWorkbook(fs.getRoot(), true);
        Biff8EncryptionKey.setCurrentUserPassword(null);*/
    }
}
