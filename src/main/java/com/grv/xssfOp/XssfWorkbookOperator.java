package com.grv.xssfOp;

import com.grv.xssfOp.xssfSheetOp.XssfSheetOperator;
import com.grv.xssfOp.xssfSheetOp.XssfSheetOperatorFactory;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

public class XssfWorkbookOperator {
    protected String filePath;
    protected File file;
    protected XSSFWorkbook workbook;
    protected XssfSheetOperatorFactory sheetOperatorFactory;

    public XssfWorkbookOperator(String filePath) throws IOException {
        super();
        FileInputStream inputStream = null;
        try {
            this.filePath = filePath;
            this.file = new File(filePath);
            inputStream = new FileInputStream(file);
            this.workbook = new XSSFWorkbook(inputStream);
            this.sheetOperatorFactory = new XssfSheetOperatorFactory(workbook);
        } catch (IOException e) {
            throw new IOException("Failed to instantiate XssfWorkBookOperator with filePath: " + filePath, e);
        } finally {
            assert inputStream != null;
            inputStream.close();
        }
    }

    public XssfSheetOperator getSheetOperator(String sheetName) {
        return sheetOperatorFactory.getSheetOperator(sheetName);
    }

    public List<XssfSheetOperator> getAllSheetOperators() {
        List<XssfSheetOperator> sheetOperators = new ArrayList<>();
        workbook.iterator().forEachRemaining(s -> sheetOperators.add(getSheetOperator(s.getSheetName())));
        return sheetOperators;
    }

    public List<XssfSheetOperator> getAllSheetOperators(List<String> exclude) {
        return getAllSheetOperators().stream().filter(op -> !exclude.contains(op.getSheetName()))
                .collect(Collectors.toList());
    }

    public XssfSheetOperator addSheet(String sheetName) {
        workbook.createSheet(sheetName);
        return getSheetOperator(sheetName);
    }

    public void addSheetAndEmptyTable(String sheetName, List<String> columnHeaders) {
        XssfSheetOperator sheetOperator = addSheet(sheetName);
		sheetOperator.addEmptyTable(columnHeaders, sheetName, 10);
    }

    public void writeChanges() throws IOException {
		try (FileOutputStream outputStream = new FileOutputStream(file)) {
			workbook.write(outputStream);
		} catch (IOException e) {
			throw e;
		} finally {
			workbook.close();
		}
    }

    public XSSFName findName(String nameName) {
        return workbook.getName(nameName);
    }
}
