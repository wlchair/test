package com.acmr.excel.model;

import java.io.Serializable;

public class RowHeight implements Serializable {
	private String excelId;
	private int rowSort;
	private int offset;

	public String getExcelId() {
		return excelId;
	}

	public void setExcelId(String excelId) {
		this.excelId = excelId;
	}

	public int getRowSort() {
		return rowSort;
	}

	public void setRowSort(int rowSort) {
		this.rowSort = rowSort;
	}

	public int getOffset() {
		return offset;
	}

	public void setOffset(int offset) {
		this.offset = offset;
	}

}
