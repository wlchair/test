package com.acmr.excel.model;

import java.io.Serializable;

public class ColWidth implements Serializable {
	private String excelId;
	private int colSort;
	private int offset;

	public String getExcelId() {
		return excelId;
	}

	public void setExcelId(String excelId) {
		this.excelId = excelId;
	}

	public int getOffset() {
		return offset;
	}

	public void setOffset(int offset) {
		this.offset = offset;
	}

	public int getColSort() {
		return colSort;
	}

	public void setColSort(int colSort) {
		this.colSort = colSort;
	}

}
