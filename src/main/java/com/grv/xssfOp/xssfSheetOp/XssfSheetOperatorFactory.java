/**
 *
 */
package com.grv.xssfOp.xssfSheetOp;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Objects;
import java.util.concurrent.ConcurrentHashMap;

/**
 * @author kgaurav
 *
 */
public class XssfSheetOperatorFactory {

    private ConcurrentHashMap<String, XssfSheetOperator> sheetNameToOperator;

    private XSSFWorkbook workbook;

    public XssfSheetOperatorFactory(XSSFWorkbook workbook) {
        super();
        this.workbook = workbook;
        init();
    }

    private void init() {
        sheetNameToOperator = new ConcurrentHashMap<String, XssfSheetOperator>(workbook.getNumberOfSheets());
    }

    public XssfSheetOperator getSheetOperator(String sheetName) {
        XssfSheetOperator sheetOperator = null;
        if (!sheetNameToOperator.containsKey(sheetName)) {
            XSSFSheet sheet = workbook.getSheet(sheetName);
            Objects.requireNonNull(sheet, "Sheet Not Found: " + sheetName);
            sheetOperator = new XssfSheetOperator(sheet);
            sheetNameToOperator.put(sheetName, sheetOperator);
        }
        return sheetNameToOperator.get(sheetName);
    }
}
