package org.mig.omalang;

import java.io.File;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.List;

import javax.xml.namespace.QName;

import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.mutable.MutableBoolean;
import org.docx4j.TextUtils;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.Body;
import org.jvnet.jaxb2_commons.ppp.Child;

public class WordDocument {
	ArrayList<Child> topChildren = new ArrayList<Child>();
	final WordprocessingMLPackage wordMlPackage;
	private String path;
	
	public WordDocument(String path) throws Docx4JException {
		this.path = path;
		File docFile = new File(path);
		wordMlPackage = WordprocessingMLPackage.load(docFile);
		fillTopChildren();
	}
	
	public void save() throws Docx4JException {
		save(this.path);
	}
	
	public void save(String pathAndFileName) throws Docx4JException {
		this.path = pathAndFileName;
		wordMlPackage.save(new File(pathAndFileName));
	}
	
	public void replace(org.docx4j.wml.P p, final String whatToReplace, final String replaceString) {
		// TODO extend to flatten the P -> (R -> Text), (R -> Text) ... (R -> Text) and to replace whatToReplace properly
		TraversalUtil.visit(p, new TraversalUtil.Callback() {
			
			@Override
			public void walkJAXBElements(Object parent) {
				List<?> children = getChildren(parent);
				if (children != null) {
					for (Object o : children) {
						o = XmlUtils.unwrap(o);
						this.apply(o);
						if (this.shouldTraverse(o)) {
							walkJAXBElements(o);
						}
					}
				}
			}
			
			@Override
			public boolean shouldTraverse(Object o) {
				return true;
			}
			
			@Override
			public List<Object> getChildren(Object o) {
				return TraversalUtil.getChildrenImpl(o);
			}
			
			@Override
			public List<Object> apply(Object o) {
				if (o instanceof org.docx4j.wml.Text) {
					org.docx4j.wml.Text wmlText = (org.docx4j.wml.Text) o;
					final String wmlTextValue = wmlText.getValue();
					if (StringUtils.containsAny(wmlTextValue, whatToReplace)) {
						wmlText.setValue(StringUtils.replace(wmlTextValue, whatToReplace, replaceString));
					}
				}
				return null;
			}
		});
	}
	
	public List<org.docx4j.wml.P> findParagraph(String subString) throws Exception {
		final ArrayList<org.docx4j.wml.P> foundPs = new ArrayList<org.docx4j.wml.P>();
		
		for (Child child : topChildren) {
			if (child instanceof org.docx4j.wml.P) {
				org.docx4j.wml.P pChild = (org.docx4j.wml.P) child;
				StringWriter sw = new StringWriter();
				TextUtils.extractText(pChild, sw);
				if (sw.toString().contains(subString)) {
					foundPs.add(pChild);
				}
			}
		}
		
		return foundPs;
	}
	
	public int getTopChildIndex(Object aTopChild) {
		int tcIdx = -1;
		
		final org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) wordMlPackage.getMainDocumentPart().getJaxbElement();
		final Body body = wmlDocumentEl.getBody();
		
		for (int i = 0; i < body.getContent().size(); ++i) {
			Object child = body.getContent().get(i);
			if (child instanceof javax.xml.bind.JAXBElement) {
				javax.xml.bind.JAXBElement jaxBChild = (javax.xml.bind.JAXBElement) child;
				child = jaxBChild.getValue();
			}
			if (aTopChild.hashCode() == child.hashCode()) {
				tcIdx = i;
				break;
			}
		}
		
		return tcIdx;
	}
	
	public void insertAfter(Object aTopChild, Object newObject) {
		final org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) wordMlPackage.getMainDocumentPart().getJaxbElement();
		Body body = wmlDocumentEl.getBody();
		int index = wmlDocumentEl.getBody().getContent().size() - 1;
		
		for (int i = 0; i < body.getContent().size(); ++i) {
			Object child = body.getContent().get(i);
			if (child instanceof javax.xml.bind.JAXBElement) {
				javax.xml.bind.JAXBElement jaxBChild = (javax.xml.bind.JAXBElement) child;
				child = jaxBChild.getValue();
			}
			if (aTopChild.hashCode() == child.hashCode()) {
				index = i;
				break;
			}
		}
		
		if (newObject instanceof org.docx4j.wml.Tbl) {
			javax.xml.bind.JAXBElement newJaxbObject = new javax.xml.bind.JAXBElement(new QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "tbl"),
					org.docx4j.wml.Tbl.class, newObject);
			wmlDocumentEl.getBody().getContent().add(index + 1, newJaxbObject);
		} else {
			wmlDocumentEl.getBody().getContent().add(index + 1, newObject);
		}
		
		this.topChildren.clear();
		fillTopChildren();
	}

	// todo remove the hack
	public static boolean toTrace = false;
	
	public List<Table> findTables(List<String> terms) {
		return findTables(terms.toArray(new String[0]));
	}
	
	public List<Table> findTables(String ... terms) {
		final ArrayList<Table> foundTbls = new ArrayList<Table>();
		
		for (Child child : topChildren) {
			if (child instanceof org.docx4j.wml.Tbl) {
				org.docx4j.wml.Tbl tblChild = (org.docx4j.wml.Tbl) child;
				final MutableBoolean containsAll = new MutableBoolean(true);
				
				for (String term : terms) {
					final MutableBoolean containsOne = new MutableBoolean(false);
					
//					TextTraverser.traverse(tblChild,
//							t -> StringUtils.contains(t.getValue(), term),
//							t -> containsOne.setValue(true));
					
					TextTraverser.traverseP(tblChild,
							p -> {
								if (toTrace) {
									System.out.println("finding '" + term + "' in '" + TextTraverser.extractTextValue(p) + "'");
								}
								return StringUtils.contains(TextTraverser.extractTextValue(p), term);
							},
							p -> {
								containsOne.setValue(true);
								}
							);
					
					
					containsAll.setValue(containsAll.booleanValue() && containsOne.booleanValue());
					if (!containsAll.booleanValue()) {
						break;
					}
				}
				
				if (containsAll.booleanValue()) {
					foundTbls.add(TableFactory.createFromWordDocument(tblChild));
				}
			}
		}
		return foundTbls;
	}
	
	private void fillTopChildren() {
		final org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) wordMlPackage.getMainDocumentPart().getJaxbElement();
		Body body = wmlDocumentEl.getBody();

//		final HashSet<String> topLevelWMLs = new HashSet<String>();
//		
//		new TraversalUtil(body, new Callback() {
//			private String indent = "";
//			private ArrayList<String> texts = new ArrayList<String>();
//
//			@Override
//			public List<Object> apply(Object o) {
//				String text = "";
//				if (o instanceof org.docx4j.wml.Text)
//					text = ((org.docx4j.wml.Text) o).getValue();
//				System.out.println(indent + o.getClass().getName() + "  \"" + text + "\"");
//				return null;
//			}
//
//			@Override
//			public boolean shouldTraverse(Object o) {
//				return true;
//			}
//
//			private void traversal(Object parent, boolean isUnderParent) {
//				indent += "    ";
//				if (!isUnderParent) {
//					topLevelWMLs.add(parent.getClass().getName());
//				}
//				
//				List<?> children = getChildren(parent);
//				if (children != null) {
//					for (Object o : children) {
//						o = XmlUtils.unwrap(o);
//						this.apply(o);
//						if (this.shouldTraverse(o)) {
//							traversal(o, true);
//						}
//					}
//				}
//				indent = indent.substring(0, indent.length() - 4);
//			}
//			
//			@Override
//			public void walkJAXBElements(Object parent) {
//				traversal(parent, false);
//			}
//
//			@Override
//			public List<Object> getChildren(Object o) {
//				return TraversalUtil.getChildrenImpl(o);
//			}
//		}
//		);

		List<?> childs = TraversalUtil.getChildrenImpl(body);
		for (Object child : childs) {
			if (child instanceof javax.xml.bind.JAXBElement) {
				javax.xml.bind.JAXBElement jaxBChild = (javax.xml.bind.JAXBElement) child;
				if (jaxBChild.getValue() instanceof org.docx4j.wml.Tbl) {
					topChildren.add((org.docx4j.wml.Tbl)jaxBChild.getValue());
					
//					System.out.println("-----");
//					StringWriter sw = new StringWriter();
//					try {
//						TextUtils.extractText(jaxBChild.getValue(), sw);
//					} catch (Exception e) {
//						// TODO Auto-generated catch block
//						e.printStackTrace();
//					}
//					System.out.println(sw.toString());
				}
			} else if (child instanceof org.docx4j.wml.P) {
				org.docx4j.wml.P pChild = (org.docx4j.wml.P) child;
				topChildren.add(pChild);
//				StringWriter sw = new StringWriter();
//				TextUtils.extractText(pChild, sw);
//				paras.add(sw.toString());
			} else {
				//
			}
		}
	}

	public ArrayList<Child> getTopChildren() {
		return topChildren;
	}

	public WordprocessingMLPackage getWordMlPackage() {
		return wordMlPackage;
	}
}
