package com.acmr.excel.model;

import java.io.Serializable;

public class Coordinate implements Serializable {

	private int endSortX;
	private int endSortY;
	private int startSortX;
	private int startSortY;

	public int getEndSortX() {
		return endSortX;
	}

	public void setEndSortX(int endSortX) {
		this.endSortX = endSortX;
	}

	public int getEndSortY() {
		return endSortY;
	}

	public void setEndSortY(int endSortY) {
		this.endSortY = endSortY;
	}

	public int getStartSortX() {
		return startSortX;
	}

	public void setStartSortX(int startSortX) {
		this.startSortX = startSortX;
	}

	public int getStartSortY() {
		return startSortY;
	}

	public void setStartSortY(int startSortY) {
		this.startSortY = startSortY;
	}

}
