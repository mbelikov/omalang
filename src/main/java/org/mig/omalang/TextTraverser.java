package org.mig.omalang;

import java.io.StringWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Predicate;

import org.docx4j.TextUtils;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;

// TODO rename
public class TextTraverser {
	public static class TraversalCallback implements TraversalUtil.Callback {
		@Override
		public void walkJAXBElements(Object parent) {
			List<?> children = getChildren(parent);
			if (children != null) {
				for (Object o : children) {
					o = XmlUtils.unwrap(o);
					
					doWith(o);
					if (this.shouldTraverse(o)) {
						walkJAXBElements(o);
					}
				}
			}
		}
		
		public void doWith(Object o) {
			// empty, nothing, overload me!
		}
		
		@Override
		public List<Object> getChildren(Object o) {
			return TraversalUtil.getChildrenImpl(o);
		}
		
		@Override
		public boolean shouldTraverse(Object o) {
			// default let's go
			return true;
		}
		
		@Override
		public List<Object> apply(Object o) {
			return null;
		}
	}
	
	public static void traverse(Object parent, Predicate<org.docx4j.wml.Text> tester, Consumer<org.docx4j.wml.Text> block) {
		TraversalUtil.visit(parent, new TraversalCallback() {
			@Override
			public void doWith(Object o) {
				if (o instanceof org.docx4j.wml.Text) {
					org.docx4j.wml.Text textObject = (org.docx4j.wml.Text) o;
					if (tester.test(textObject)) {
						block.accept(textObject);
					}
				}
			}			
		});
	}
	
	public static void traverseP(Object parent, Predicate<org.docx4j.wml.P> tester, Consumer<org.docx4j.wml.P> block) {
		TraversalUtil.visit(parent, new TraversalCallback() {
			@Override
			public void doWith(Object o) {
				if (o instanceof org.docx4j.wml.P) {
					org.docx4j.wml.P pObject = (org.docx4j.wml.P) o;
					if (tester.test(pObject)) {
						block.accept(pObject);
					}
				}
			}
		});
	}
	
	public static void traverseTables(List<org.docx4j.wml.Tbl> tables, Predicate<org.docx4j.wml.P> tester, Consumer<org.docx4j.wml.P> block) {
		for (org.docx4j.wml.Tbl table : tables) {
			traverseTable(table, tester, block);
		}
	}
	
	public static void traverseTable(org.docx4j.wml.Tbl table, Predicate<org.docx4j.wml.P> tester, Consumer<org.docx4j.wml.P> block) {
		TraversalUtil.visit(table, new TraversalCallback() {
			@Override
			public void doWith(Object o) {
				if (o instanceof org.docx4j.wml.P) {
					org.docx4j.wml.P textObject = (org.docx4j.wml.P) o;
					if (tester.test(textObject)) {
						block.accept(textObject);
					}
				}
			}
		});
	}
	
	public static List<String> getTableColumn(org.docx4j.wml.Tbl table, int column) {
		final List<String> columnStrings = new ArrayList<String>();
		TableFactory.createFromWordDocument(table).
				process(c -> c.c == column, c -> columnStrings.add(c.value));
		return columnStrings;
	}
	
	public static String extractTextValue(Object obj) {
		final StringWriter sw = new StringWriter();
		try {
			TextUtils.extractText(obj, sw);
		} catch (Exception e) {}
		return sw.toString();
	}
	
	public static List<String> getTableRow(org.docx4j.wml.Tbl table, int row) {
		final List<String> rowStrings = new ArrayList<String>();
		TableFactory.createFromWordDocument(table).
			process(c -> c.r == row, c -> rowStrings.add(c.value));
		return rowStrings;
	}
	
	public static org.docx4j.wml.P createP(String newValue) {
		org.docx4j.wml.P p = new org.docx4j.wml.P();
		setTextToP(p, newValue);
		return p;
	}
	
	public static void setTextToP(org.docx4j.wml.P p, String newValue) {
		org.docx4j.wml.R r;
		
		// TODO: quick & dirty hack
		if (p.getContent().size() > 0 && p.getContent().get(0) instanceof org.docx4j.wml.R) {
			r = (org.docx4j.wml.R)p.getContent().get(0);
		} else if (p.getContent().size() > 0 && XmlUtils.unwrap(p.getContent().get(0)) instanceof org.docx4j.wml.ContentAccessor
				&& ((org.docx4j.wml.ContentAccessor)XmlUtils.unwrap(p.getContent().get(0))).getContent().size() > 0
				&& ((org.docx4j.wml.ContentAccessor)XmlUtils.unwrap(p.getContent().get(0))).getContent().get(0) instanceof org.docx4j.wml.R) {
			r = (org.docx4j.wml.R)((org.docx4j.wml.ContentAccessor)XmlUtils.unwrap(p.getContent().get(0))).getContent().get(0);
		} else {
			r = new org.docx4j.wml.R();
			org.docx4j.wml.Text text = new org.docx4j.wml.Text(); 
			r.getContent().add(text);
		}
		p.getContent().clear();
		p.getContent().add(r);
		
		org.docx4j.wml.Text text = null;
		
		// one more dirty hack
		if (XmlUtils.unwrap(r.getContent().get(0)) instanceof org.docx4j.wml.Text) {
			text = (org.docx4j.wml.Text)XmlUtils.unwrap(r.getContent().get(0));
//		} else {
		} else if (r.getContent().size() > 1) {
			text = (org.docx4j.wml.Text)XmlUtils.unwrap(r.getContent().get(1));
		} else {
			text = new org.docx4j.wml.Text(); 
			r.getContent().add(text);
		}
		
		r.getContent().clear();
		r.getContent().add(text);
		text.setValue(newValue);
	}
	
	public static List<String> compare(List<String> left, List<String> right) {
		final List<String> diffList = new ArrayList(left);
		diffList.removeAll(right);
		return diffList;
	}
	
	public static List<String> subList(List<String> list, List<Integer> indexes) {
		final List<String> subLst = new ArrayList<String>();
		for (int i : indexes) {
			subLst.add(list.get(i));
		}
		return subLst;
	}
}
