package org.mig.omalang.model;

public abstract class Document implements Formattable {
	private String fileName;
	private Matrix<Content> internalModel;
}
