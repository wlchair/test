package com.acmr.excel.model.copy;

import java.io.Serializable;

public class Target implements Serializable {
	private int colSort;
	private int rowSort;

	public int getColSort() {
		return colSort;
	}

	public void setColSort(int colSort) {
		this.colSort = colSort;
	}

	public int getRowSort() {
		return rowSort;
	}

	public void setRowSort(int rowSort) {
		this.rowSort = rowSort;
	}

}
