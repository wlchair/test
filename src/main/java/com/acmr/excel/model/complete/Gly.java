package com.acmr.excel.model.complete;

import java.io.Serializable;

public class Gly implements Serializable {
	private String aliasY;
	private int top;
	private int height;
	private int index;
	private OperProp operProp = new OperProp();

	public String getAliasY() {
		return aliasY;
	}

	public void setAliasY(String aliasY) {
		this.aliasY = aliasY;
	}

	public int getTop() {
		return top;
	}

	public void setTop(int top) {
		this.top = top;
	}

	public int getHeight() {
		return height;
	}

	public void setHeight(int height) {
		this.height = height;
	}

	public int getIndex() {
		return index;
	}

	public void setIndex(int index) {
		this.index = index;
	}

	public OperProp getOperProp() {
		return operProp;
	}

	public void setOperProp(OperProp operProp) {
		this.operProp = operProp;
	}

}
