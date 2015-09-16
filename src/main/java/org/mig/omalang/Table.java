package org.mig.omalang;

import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;

import javax.xml.bind.JAXBException;

import org.apache.commons.lang.StringUtils;
import org.docx4j.XmlUtils;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.xlsx4j.sml.CTRst;
import org.xlsx4j.sml.CTXstringWhitespace;
import org.xlsx4j.sml.STCellType;

public class Table implements Comparable<Table> {
	private org.docx4j.wml.Tbl wmlTbl;
	private List<org.xlsx4j.sml.Row> excellRows;
	private final ArrayList<ArrayList<Cell>> tableModel;
	private int rowCount;
	private int columnCount;
	
	public static class Cell {
		public String value;
		public int r;
		public int c;
		private String index;
		private org.docx4j.wml.Tc wordInternal;
		private org.xlsx4j.sml.Cell excelInternal;
		
		public Cell(String value, String index, int r, int c) {
			super();
			this.value = value;
			this.index = index;
			this.r = r;
			this.c = c;
		}

		public org.docx4j.wml.Tc getWordInternal() {
			return wordInternal;
		}

		public Cell withWordInternal(org.docx4j.wml.Tc wordCell) {
			this.wordInternal = wordCell;
			return this;
		}

		public org.xlsx4j.sml.Cell getExcelInternal() {
			return excelInternal;
		}

		public Cell withExcelInternal(org.xlsx4j.sml.Cell excelCell) {
			this.excelInternal = excelCell;
			return this;
		}

		public String getValue() {
			return value;
		}

		public void setValue(String value) {
			this.value = value;
			changeModel(this.value);
		}

		private void changeModel(String newValue) {
			if (wordInternal != null) {
				setTextToWordTc(wordInternal, newValue);
			} else if (excelInternal != null) {
				setTextToExcellTc(excelInternal, newValue);
			}
		}

		public String getIndex() {
			return index;
		}
	}
	
	Table(ArrayList<ArrayList<Cell>> tableModel, int rowCount,
			int columnCount, Tbl wmlTbl) {
		super();
		this.tableModel = tableModel;
		this.rowCount = rowCount;
		this.columnCount = columnCount;
		this.wmlTbl = wmlTbl;
	}

	Table(ArrayList<ArrayList<Cell>> tableModel, int rowCount,
			int columnCount) {
		super();
		this.tableModel = tableModel;
		this.rowCount = rowCount;
		this.columnCount = columnCount;
	}
	
	public static void setTextToWordTc(Tc wordInternal, String newValue) {
		org.docx4j.wml.P p = (org.docx4j.wml.P)wordInternal.getContent().get(0);
		wordInternal.getContent().clear();
		
		TextTraverser.setTextToP(p, newValue);
		
		wordInternal.getContent().add(p);
	}

	private static void setTextToExcellTc(org.xlsx4j.sml.Cell excelInternal, String newValue) {
		// TODO Auto-generated method stub
		CTRst ctrst = new CTRst();
		CTXstringWhitespace ctxString = new CTXstringWhitespace();
		ctxString.setValue(newValue);
		ctrst.setT(ctxString);
		
		excelInternal.setT(STCellType.INLINE_STR);
		excelInternal.setIs(ctrst);
	}

	public org.docx4j.wml.Tbl getWmlTbl() {
		return wmlTbl;
	}

//	public void setWmlTbl(org.docx4j.wml.Tbl wmlTbl) {
//		this.wmlTbl = wmlTbl;
//	}

	public String getWholeText(int row, int col) {
		return tableModel.get(row).get(col).getValue();
	}

	public List<Cell> getTableRow(int row) {
		return tableModel.get(row);
	}
	
	public List<Cell> getTableColumn(int column) {
		final List<Cell> columnStrings = new ArrayList<Cell>();
		for (ArrayList<Cell> row : tableModel) {
			columnStrings.add(row.get(column));
		}
		return columnStrings;
	}
	
	public void process(Predicate<Cell> tester, Consumer<Cell> block) {
		for (List<Cell> row : tableModel) {
			for (Cell cell : row) {
				if (tester.test(cell)) {
					block.accept(cell);
				}
			}
		}
	}
	
	public void process(Predicate<List<Cell>> tester, Function<List<Cell>, Cell> mapper, Consumer<Cell> block) {
		for (List<Cell> row : tableModel) {
			if (tester.test(row)) {
				Cell cell = mapper.apply(row);
				block.accept(cell);
			}
		}
	}
	
	public void processRow(Predicate<List<Cell>> tester, Consumer<List<Cell>> block) {
		for (List<Cell> row : tableModel) {
			if (tester.test(row)) {
				block.accept(row);
			}
		}
	}
	
	public boolean contains(String ... terms) {
		// todo
		List<?> tblObjs = wmlTbl.getContent();
		for (Object tblObj : tblObjs) {
			org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) tblObj;
			
		}
		
		return false;
	}

	@Override
	public int compareTo(Table o) {
		if (this.rowCount != o.rowCount || this.columnCount != o.columnCount) {
			if (this.rowCount < o.rowCount && this.columnCount < o.columnCount) {
				return -1;
			}
			return 1;
		}
		
		boolean equals = true;
		for (int i = 0; i < rowCount; ++i) {
			for (int j = 0; j < columnCount; ++j) {
				equals &= this.getWholeText(i, j).equals(o.getWholeText(i, j)); 
			}
		}
		
		if (!equals) {
			return 1;
		}
		
		return 0;
	}

	public void useAsColumn(int i, List<String> strColumn) throws JAXBException {
		if (strColumn.size() == 0) {
			return;
		}
		
		// TODO tbl model, only word implementation
		// todo use the existing rows and add new if needed
		List<Object> tblObjs = wmlTbl.getContent();
		org.docx4j.wml.Tr headerTr = (org.docx4j.wml.Tr) tblObjs.get(0);
		org.docx4j.wml.Tr etalonTr = Utils.cloneDocx4jObject(
				(org.docx4j.wml.Tr) tblObjs.get(1));
		
		TextTraverser.traverse(etalonTr,
				t -> StringUtils.isNotEmpty(t.getValue()), t -> t.setValue(""));
		
		final int extensionCount = strColumn.size() - tblObjs.size() + 1;
		for (int iter = 0; iter < extensionCount; ++iter) {
			tblObjs.add(Utils.cloneDocx4jObject(etalonTr));
		}
		
		for (int iter = 1; iter < tblObjs.size(); ++iter) {
			org.docx4j.wml.Tr aTr = (org.docx4j.wml.Tr) tblObjs.get(iter);
			List<Object> tcObjs = aTr.getContent();
			if (strColumn.size() > (iter -1)) {
				setTextToWordTc((org.docx4j.wml.Tc)XmlUtils.unwrap(tcObjs.get(i)), strColumn.get(iter - 1));
			}
		}
	}
}
