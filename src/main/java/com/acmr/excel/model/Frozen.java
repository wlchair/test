package com.acmr.excel.model;

import java.io.Serializable;

public class Frozen implements Serializable {
	private int frozenSortX;
	private int frozenSortY;
	private int startSortX;
	private int startSortY;

	public int getFrozenSortX() {
		return frozenSortX;
	}

	public void setFrozenSortX(int frozenSortX) {
		this.frozenSortX = frozenSortX;
	}

	public int getFrozenSortY() {
		return frozenSortY;
	}

	public void setFrozenSortY(int frozenSortY) {
		this.frozenSortY = frozenSortY;
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
