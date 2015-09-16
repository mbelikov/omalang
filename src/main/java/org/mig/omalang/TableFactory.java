package org.mig.omalang;

import java.io.StringWriter;
import java.util.ArrayList;
import java.util.List;

import org.docx4j.TextUtils;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.SpreadsheetML.SharedStrings;
import org.mig.omalang.Table.Cell;
import org.xlsx4j.exceptions.Xlsx4jException;
import org.xlsx4j.jaxb.Context;
import org.xlsx4j.sml.CTRst;
import org.xlsx4j.sml.CTXstringWhitespace;
import org.xlsx4j.sml.Row;
import org.xlsx4j.sml.STCellType;

public class TableFactory {
	public static Table createFromWordDocument(org.docx4j.wml.Tbl wmlTbl) {
		final ArrayList<ArrayList<Cell>> tableModel = new ArrayList<ArrayList<Cell>>();
		
		final List<?> trObjList = wmlTbl.getContent();
		
		int rowCount = trObjList.size();
		int columnCount = 0;
		for (Object trObj : trObjList) {
			final org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) trObj;
			if (tr.getContent().size() > columnCount) {
				columnCount = tr.getContent().size();
			}
		}
		
		String[][] mergeStoreVals = new String[rowCount][columnCount];
		
		for (int i = 0; i < trObjList.size(); ++i) {
			
			final org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) trObjList.get(i);
			final ArrayList<Cell> rowStrings = new ArrayList<Cell>();
			tableModel.add(rowStrings);
			
			final List<?> jaxbTcObjList = tr.getContent();
			
			for (int j = 0; j < jaxbTcObjList.size(); ++j) {
				final org.docx4j.wml.Tc tc = (org.docx4j.wml.Tc) XmlUtils.unwrap(jaxbTcObjList.get(j));
				String tcTxtVal = TextTraverser.extractTextValue(tc);
				
				// vmerge
				if (tc.getTcPr().getVMerge() != null) {
					if (tc.getTcPr().getVMerge().getVal() != null) {
						// v merge start
						mergeStoreVals[i][j] = tcTxtVal;
					} else {
						// to take value from above
						mergeStoreVals[i][j] = mergeStoreVals[i-1][j];
						tcTxtVal = "vmerge: " + mergeStoreVals[i-1][j];
					}
				}
				
				rowStrings.add(new Cell(tcTxtVal, "r:" + i + "c:" + j, i, j).withWordInternal(tc));
				
				// hmerge
				if (tc.getTcPr().getGridSpan() != null) {
					final int spanNum = tc.getTcPr().getGridSpan().getVal().intValue();
					for (int sc = 1; sc < spanNum; ++sc) {
						
						rowStrings.add(
								new Cell("hmerge: " + tcTxtVal, "r:" + i + "c:" + (j + sc), i, j + sc).withWordInternal(tc));
					}
				}
			}
		}

		Table createdTable = new Table(tableModel, rowCount, columnCount, wmlTbl);
//		createdTable.setWmlTbl(wmlTbl);
		return createdTable;
	}
	
	public static Table createFromExcelDocument(ExcelDocument ed, int worksheetNum) throws Exception {
		final ArrayList<ArrayList<Cell>> tableModel = new ArrayList<ArrayList<Cell>>();
		
		final SharedStrings sharedStrings = (SharedStrings) ed.getSpreadsheetMLPackage().getParts().get(new PartName("/xl/sharedStrings.xml"));
		final List<CTRst> sis = sharedStrings.getContents().getSi();
		
		final List<Row> rows = ed.getSpreadsheetMLPackage().getWorkbookPart().getWorksheet(worksheetNum).getContents().getSheetData().getRow();
		
		final int rowCount = rows.size();
		int columnCount = 0;
		
		int i = 0;
		
		for (Row row : rows) {
			final List<org.xlsx4j.sml.Cell> cells = row.getC();
			final ArrayList<Cell> rowModel = new ArrayList<Cell>();
			tableModel.add(rowModel);
			
			int j = 0;
			for (org.xlsx4j.sml.Cell cell : cells) {
				String value;
				
				if (STCellType.S.equals(cell.getT())) {
					CTRst ctrst = sis.get(Integer.parseInt(cell.getV()));
					
					final StringWriter sw = new StringWriter();
					TextUtils.extractText(ctrst, sw, 
							Context.jcSML,
							"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
							"CT_Rst", CTRst.class);
					value = sw.getBuffer().toString();
					
//					CTXstringWhitespace whiteSpace = sis.get(Integer.parseInt(cell.getV())).getT();
//					if (whiteSpace != null) {
//						value = whiteSpace.getValue();
//					} else  {
//						System.out.println(cell);
//						value = "null";
//					}
				} else {
					value = cell.getV();
				}
				
				rowModel.add(new Cell(value, cell.getR(), i, j++).withExcelInternal(cell));
			}
			
			++i;
			
			if (rowModel.size() >= columnCount) {
				columnCount = rowModel.size();
			} else {
				System.out.println("another column count: " + rowModel.size());
			}
		}
	
		return new Table(tableModel, rowCount, columnCount);
	}
}
