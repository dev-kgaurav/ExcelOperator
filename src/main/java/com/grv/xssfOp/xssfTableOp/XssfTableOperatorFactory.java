/**
 *
 */
package com.grv.xssfOp.xssfTableOp;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;

import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ConcurrentHashMap;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * @author kgaurav
 *
 */
public class XssfTableOperatorFactory {
    protected ConcurrentHashMap<String, XssfTableOperator> tableNameToOperator;
    private XSSFSheet sheet;
    private Map<String, XSSFTable> tableNametoTable;

    public XssfTableOperatorFactory(XSSFSheet sheet) {
        super();
        this.sheet = sheet;
        init();
    }

    private void init() {
        tableNameToOperator = new ConcurrentHashMap<String, XssfTableOperator>(sheet.getTables().size());
        updateLocalTableList();
    }

    private void updateLocalTableList() {
        tableNametoTable = sheet.getTables().stream().collect(Collectors.toMap(XSSFTable::getName, Function.identity()));
    }

    protected XSSFTable findTable(String tableName) {
        if (!tableNametoTable.containsKey(tableName))
            updateLocalTableList();

        XSSFTable table = tableNametoTable.get(tableName);
        Objects.requireNonNull(table, "Table not found with name :" + tableName);
        return table;
    }

    public XssfTableOperator getTableOperator(String tableName) {
        XssfTableOperator tableOperator = null;

        if (!tableNameToOperator.contains(tableName)) {
            tableOperator = new XssfTableOperator(findTable(tableName));
            tableNameToOperator.put(tableName, tableOperator);
        }
        return tableNameToOperator.get(tableName);
    }
}
