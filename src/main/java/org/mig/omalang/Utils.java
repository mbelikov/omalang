package org.mig.omalang;

import java.util.ArrayList;
import java.util.List;

import javax.xml.bind.JAXBException;

import org.docx4j.XmlUtils;

public class Utils {
	@SuppressWarnings({ "restriction", "unchecked" })
	public static <T> T cloneDocx4jObject(T object) throws JAXBException {
		return (T) XmlUtils.unmarshalString(XmlUtils.marshaltoString(object));
	}
	
	public static <T> List<T> distinct(List<T> vector) {
		List<T> distinctVector = new ArrayList<T>();
		vector.stream().distinct().forEach(s -> distinctVector.add(s));
		return distinctVector;
	}
}
