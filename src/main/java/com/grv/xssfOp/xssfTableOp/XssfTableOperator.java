/**
 *
 */
package com.grv.xssfOp.xssfTableOp;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;

import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * @author kgaurav
 *
 */
public class XssfTableOperator {

    private Map<String, CTTableColumn> headerNameToColumn;

    private XSSFTable table;
    private XSSFSheet sheet;
    private XSSFWorkbook wb;
    private int firstRow;
    private int lastRow;


    public XssfTableOperator(XSSFTable table) {
        super();
        this.table = table;
        init();
    }

    private void init() {
        sheet = table.getXSSFSheet();
        wb = sheet.getWorkbook();
        firstRow = table.getStartRowIndex();
        lastRow = table.getEndRowIndex();

        headerNameToColumn = table.getCTTable().getTableColumns().getTableColumnList().stream().collect(Collectors.toMap(CTTableColumn::getName, Function.identity()));
    }

    protected CTTableColumn getTableColumn(String columnHeaderName) {
        return headerNameToColumn.get(columnHeaderName);
    }

    public void addDropDownValidationToColumn(String columnName, XSSFName dataSourceNamedRange) {
        int colIndex = table.findColumnIndex(columnName);

        Validate.isTrue(colIndex >= 0, "Column not found with Column-Name: %s", columnName);

        CellRangeAddressList regions = new CellRangeAddressList(firstRow + 1, lastRow, colIndex, colIndex);

        XSSFDataValidationHelper dataValidationHelper = (XSSFDataValidationHelper) sheet.getDataValidationHelper();
        XSSFDataValidationConstraint constraint = (XSSFDataValidationConstraint) dataValidationHelper
                .createFormulaListConstraint(dataSourceNamedRange.getNameName());

        XSSFDataValidation dataValidation = (XSSFDataValidation) dataValidationHelper.createValidation(constraint, regions);

        sheet.getDataValidations();

        dataValidation.setEmptyCellAllowed(true);
        dataValidation.setSuppressDropDownArrow(true);
        sheet.addValidationData(dataValidation);
    }

    public XSSFName addDefinedNameForColumn(String columnHeaderName) {

        String definedName = generateColumnNamedRangeName(columnHeaderName);
        String referToFormula = getColumnReferenceFormula(columnHeaderName);

        Predicate<XSSFName> definedNameMatcher = n -> n.getSheetIndex() == -1 && n.getRefersToFormula().equals(referToFormula);

        boolean isDefinedNameExist = wb.getNames(definedName).stream().anyMatch(definedNameMatcher);

        if (isDefinedNameExist)
            return wb.getNames(definedName).stream().filter(definedNameMatcher).findFirst().get();

        XSSFName definedXssfName = wb.createName();
        definedXssfName.setNameName(definedName);
        definedXssfName.setRefersToFormula(referToFormula);
        definedXssfName.setComment(String.format("Range of values in column: [%s] of [%s] table", columnHeaderName, table.getDisplayName()));
        return definedXssfName;
    }

    public XSSFName setCellValueByColumn(String columnHeaderName, List<String> data, boolean doDefineName) {
        XSSFRow row = null;

        CellReference columnFirstCellReference = getColumnFirstCellReference(columnHeaderName);

        int sheetColIndex = columnFirstCellReference.getCol();
        int startRowIndex = columnFirstCellReference.getRow();

        data.add(0, columnHeaderName);

        for (Iterator<String> iterator = data.iterator(); iterator.hasNext(); ) {
            row = fetchRow(startRowIndex++);
            row.createCell(sheetColIndex).setCellValue(iterator.next());
        }

        if (doDefineName)
            return addDefinedNameForColumn(columnHeaderName);

        return null;
    }

    public void setCellValueByRow(int rowIndex, List<String> data) {
        int startColIndex = table.getStartColIndex();
        int EndColIndex = table.getEndColIndex();

        XSSFRow row = fetchRow(rowIndex);

        Iterator<String> dataIterator = data.iterator();

        IntStream.rangeClosed(startColIndex, EndColIndex).mapToObj(i -> row.createCell(i))
                .forEach(c -> c.setCellValue(dataIterator.next()));
    }

    public void setColumnNameOrdered(List<String> columnNames) {
        Iterator<String> itr = columnNames.iterator();
        table.getColumns().forEach(c -> c.setName(itr.next()));
    }

    public CTTableStyleInfo getTableStyleInfo() {
        CTTableStyleInfo ctTableStyleInfo = null;
        CTTable ctTable = table.getCTTable();

        if (table.getStyle() == null)
            ctTableStyleInfo = ctTable.addNewTableStyleInfo();
        else
            ctTableStyleInfo = ctTable.getTableStyleInfo();

        return ctTableStyleInfo;
    }

    public void setTableStyle(String tableStyleName) {
        CTTableStyleInfo ctTableStyleInfo = getTableStyleInfo();

        ctTableStyleInfo.setName(tableStyleName);
        ctTableStyleInfo.setShowColumnStripes(false);
        ctTableStyleInfo.setShowRowStripes(false);
    }

    private XSSFRow fetchRow(int rowIndex) {
        if (sheet.getRow(rowIndex) == null)
            return sheet.createRow(rowIndex);
        else
            return sheet.getRow(rowIndex);
    }

    private CellReference getColumnFirstCellReference(String columnHeaderName) {
        CellReference firstCellReference = table.getCellReferences().getFirstCell();

        int tableColIndex = table.findColumnIndex(columnHeaderName);

        int sheetColIndex = firstCellReference.getCol() + tableColIndex;
        int sheetRowIndex = firstCellReference.getRow();

        return new CellReference(sheetRowIndex, sheetColIndex);
    }

    public String generateColumnNamedRangeName(String columnHeaderName) {
        return String.format("%s_%s_%s", "RNG", sheet.getSheetName(), StringUtils.replace(columnHeaderName, " ", "_")).toUpperCase();
    }

    private String getColumnReferenceFormula(String columnHeaderName) {
        return String.format("%s[%s]", table.getName(), columnHeaderName);
    }
}
