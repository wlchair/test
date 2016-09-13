package com.acmr.excel.util;

import com.acmr.helper.model.JSONReturnData;

/**
 * 自定义返回json类
 *
 */
public class JsonReturn extends JSONReturnData {
	private static final long serialVersionUID = 1L;
	/**
	 * 行号
	 */
	private int rowNum;
	/**
	 * 列号
	 */
	private int colNum;
	
	/**
	 * 总像素
	 */
	private int rowLength;
	/**
	 * 
	 * @return
	 */
	private int maxPixel;

	private String displayRowStartAlias ;
	private String displayColStartAlias ;
	private int dataRowStartIndex ;
	private int dataColStartIndex ;

	// private Object returnParam;

	public int getMaxPixel() {
		return maxPixel;
	}

	public void setMaxPixel(int maxPixel) {
		this.maxPixel = maxPixel;
	}

	public int getRowLength() {
		return rowLength;
	}

	public void setRowLength(int rowLength) {
		this.rowLength = rowLength;
	}

	public int getRowNum() {
		return rowNum;
	}

	public void setRowNum(int rowNum) {
		this.rowNum = rowNum;
	}

	public int getColNum() {
		return colNum;
	}

	public void setColNum(int colNum) {
		this.colNum = colNum;
	}

	public JsonReturn(int code1, String msg1) {
		super(code1, msg1);
	}

	public JsonReturn(Object data1) {
		super(data1);
	}

	public String getDisplayRowStartAlias() {
		return displayRowStartAlias;
	}

	public void setDisplayRowStartAlias(String displayRowStartAlias) {
		this.displayRowStartAlias = displayRowStartAlias;
	}

	public String getDisplayColStartAlias() {
		return displayColStartAlias;
	}

	public void setDisplayColStartAlias(String displayColStartAlias) {
		this.displayColStartAlias = displayColStartAlias;
	}

	public int getDataRowStartIndex() {
		return dataRowStartIndex;
	}

	public void setDataRowStartIndex(int dataRowStartIndex) {
		this.dataRowStartIndex = dataRowStartIndex;
	}

	public int getDataColStartIndex() {
		return dataColStartIndex;
	}

	public void setDataColStartIndex(int dataColStartIndex) {
		this.dataColStartIndex = dataColStartIndex;
	}

	
}
