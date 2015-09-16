package org.mig.omalang;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;

public class ExcelDocument {
	private final SpreadsheetMLPackage spreadsheetMLPackage;
	private final Table table;
	private String path;

	public ExcelDocument(String path) throws Exception {
		File docFile = new File(path);
		spreadsheetMLPackage = SpreadsheetMLPackage.load(docFile);
		// todo do with 2
		table = TableFactory.createFromExcelDocument(this, 2);
		this.path = path;
	}

	public SpreadsheetMLPackage getSpreadsheetMLPackage() {
		return spreadsheetMLPackage;
	}

	public List<String> getColumn(int column) {
		final List<String> columnStrings = new ArrayList<String>();
		
		table.process(c -> c.c == column, c -> columnStrings.add(c.value));
		
		return columnStrings;
	}
	
	public List<String> getFilteredColumn(
			Predicate<List<org.mig.omalang.Table.Cell>> tester,
			Function<List<org.mig.omalang.Table.Cell>, org.mig.omalang.Table.Cell> mapper,
			Consumer<org.mig.omalang.Table.Cell> block
			) {
		table.process(tester, mapper, block);
		return null;
	}
	
	public void processRow(Predicate<List<org.mig.omalang.Table.Cell>> tester, Consumer<List<org.mig.omalang.Table.Cell>> block) {
		table.processRow(tester, block);
	}

	public Table getTable() {
		return table;
	}
	
	public void save() throws Docx4JException {
		save(this.path);
	}
	
	public void save(String path) throws Docx4JException {
		this.path = path;
		spreadsheetMLPackage.save(new File(this.path));
	}
	
	public void process(Predicate<Table.Cell> tester, Consumer<Table.Cell> block) {
		table.process(tester, block);
	}
	
	public String getCellValue(int r, int c) {
		return table.getWholeText(r, c);
	}
	
//	public void setCellValue(int r, int c, String value) {
//		// TODO
//	}
}
