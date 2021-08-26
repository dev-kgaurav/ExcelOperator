package com.grv.xssfOp.xssfSheetOp;

import com.grv.xssfOp.xssfTableOp.XssfTableOperator;
import com.grv.xssfOp.xssfTableOp.XssfTableOperatorFactory;
import org.apache.commons.lang3.EnumUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class XssfSheetOperator {

    private static final String DEFAULT_TABLE_STYLE = "TableStyleMedium13";
    private static Function<String, String> toTitleCase = s -> Stream.of(StringUtils.splitByCharacterTypeCamelCase(s))
            .map(StringUtils::capitalize).collect(Collectors.joining(" "));
    protected XSSFSheet sheet;
    private XssfTableOperatorFactory tableOperatorFactory;

    public XssfSheetOperator(XSSFSheet sheet) {
        super();
        this.sheet = sheet;
        this.tableOperatorFactory = new XssfTableOperatorFactory(sheet);
    }

    public static String generateTableName(String simpleName) {
        return "TBL_" + StringUtils.replace(simpleName, " ", "_").toUpperCase();
    }

    public static String generateSheetName(String camelCaseString) {
        return StringUtils.replace(toTitleCase.apply(camelCaseString), " ", "_").toUpperCase();
    }

    public static String formatColumnName(String camleCaseColName) {
        return toTitleCase.apply(camleCaseColName);
    }

    public XssfTableOperator geTableOperator(String tableName) {
        return tableOperatorFactory.getTableOperator(tableName);
    }

    public List<XssfTableOperator> getAllTableOperators() {
        return sheet.getTables().stream().map(XssfTableOperator::new).collect(Collectors.toList());
    }

    public XSSFSheet getSheet() {
        return sheet;
    }

    public String getSheetName() {
        return sheet.getSheetName();
    }

    private List<String> getStringtRowCellValues(int rowIndex) {
        List<String> columnNames = new ArrayList<String>();
        sheet.getRow(rowIndex).forEach(c -> columnNames.add(c.getStringCellValue()));
        return columnNames;
    }

    public <E extends Enum<E>> XSSFTable addEnumValueRangeTable(Class<E> enumType, Function<E, String> extractor) {
        List<String> dataList = EnumUtils.getEnumList(enumType).stream().map(extractor).collect(Collectors.toList());
        return addNamedValueRangeTable(dataList, enumType.getSimpleName());
    }

    public XssfTableOperator createTableOnSheetWithData(String startCellRef) {

        XSSFTable newTable = sheet.createTable(estimateAreaReferenceToCreateTable(startCellRef));

        String newName = generateTableName(getSheetName());

        newTable.setName(newName);
        newTable.setDisplayName(newName);

        XssfTableOperator tableOperator = tableOperatorFactory.getTableOperator(newName);

        tableOperator.setColumnNameOrdered(getStringtRowCellValues(newTable.getStartRowIndex()));
        tableOperator.setTableStyle(DEFAULT_TABLE_STYLE);

        return tableOperator;
    }

    public <E extends Enum<E>> XSSFTable addNamedValueRangeTable(List<String> columnDataList, String columnHeaderName) {
        columnDataList.sort(String::compareToIgnoreCase);

        int dataRowCount = columnDataList.size();
        int dataColCount = 1;

        AreaReference newTableAreaReference = calculateAreaReferenceForNewTable(dataRowCount, dataColCount);
        String onlyColumnName = formatColumnName(columnHeaderName);

        XSSFTable newTable = addNewTable(newTableAreaReference, columnHeaderName);

        newTable.getColumns().get(0).setName(onlyColumnName);

        XssfTableOperator tableOperator = tableOperatorFactory.getTableOperator(newTable.getName());

        tableOperator.setCellValueByColumn(onlyColumnName, columnDataList, true);
        tableOperator.setTableStyle(DEFAULT_TABLE_STYLE);

        return newTable;
    }

    public XSSFTable addEmptyTable(List<String> columnHeaders, String indicativeTableName, int numberOfRows) {
        int dataRowCount = numberOfRows;
        int dataColCount = columnHeaders.size();

        AreaReference newTableAreaReference = calculateAreaReferenceForNewTable(dataRowCount, dataColCount);

        XSSFTable table = addNewTable(newTableAreaReference, indicativeTableName);

        XssfTableOperator tableOperator = tableOperatorFactory.getTableOperator(table.getName());

        tableOperator.setColumnNameOrdered(columnHeaders);

        tableOperator.setCellValueByRow(table.getStartRowIndex(), columnHeaders);

        tableOperator.setTableStyle(DEFAULT_TABLE_STYLE);

        return table;

    }

    private XSSFTable addNewTable(AreaReference newTableAreaReference, String indicativeName) {
        String actualTableName = generateTableName(indicativeName);
        XSSFTable newTable = sheet.createTable(newTableAreaReference);
        newTable.setName(actualTableName);
        newTable.setDisplayName(actualTableName);
        return newTable;
    }

    private AreaReference estimateAreaReferenceToCreateTable(String startCellRef) {
        CellReference topLeft = new CellReference(startCellRef);

        int lastRowIndex = sheet.getLastRowNum();
        int lastColIndex = topLeft.getCol() + sheet.getRow(topLeft.getRow()).getLastCellNum() - 1;

        CellReference botRight = new CellReference(lastRowIndex, lastColIndex);

        return new AreaReference(topLeft, botRight, SpreadsheetVersion.EXCEL2007);
    }

    private AreaReference calculateAreaReferenceForNewTable(int dataRowCount, int dataColCount) {
        AtomicInteger tableStartColIndex = new AtomicInteger(0);
        AtomicInteger tableStartRowIndex = new AtomicInteger(1);

        sheet.getTables().stream().forEach(tbl -> {
            if (tbl.getEndColIndex() >= tableStartColIndex.intValue()) {
                tableStartColIndex.set(tbl.getEndColIndex() + 2);
                tableStartRowIndex.set(tbl.getStartRowIndex());
            }
        });

        CellReference topLeft = new CellReference(tableStartRowIndex.get(), tableStartColIndex.get());
        CellReference botRight = new CellReference(tableStartRowIndex.get() + dataRowCount,
                tableStartColIndex.get() + dataColCount - 1);

        return new AreaReference(topLeft, botRight, SpreadsheetVersion.EXCEL2007);
    }

    public void disableGridLines() {
        sheet.setDisplayGridlines(false);
    }
}
