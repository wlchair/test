package com.acmr.excel.model.copy;

import java.io.Serializable;

public class Orignal implements Serializable {
	private int endColSort;
	private int endRowSort;
	private int startColSort;
	private int startRowSort;

	public int getEndColSort() {
		return endColSort;
	}

	public void setEndColSort(int endColSort) {
		this.endColSort = endColSort;
	}

	public int getEndRowSort() {
		return endRowSort;
	}

	public void setEndRowSort(int endRowSort) {
		this.endRowSort = endRowSort;
	}

	public int getStartColSort() {
		return startColSort;
	}

	public void setStartColSort(int startColSort) {
		this.startColSort = startColSort;
	}

	public int getStartRowSort() {
		return startRowSort;
	}

	public void setStartRowSort(int startRowSort) {
		this.startRowSort = startRowSort;
	}

}
