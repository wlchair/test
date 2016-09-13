package com.acmr.rmi.service;

import acmr.excel.pojo.ExcelBook;

public interface RmiService {
	/**
	 * 保存excel
	 * @param excelId
	 * @param excelBook
	 */
	public void saveExcelBook(String excelId,ExcelBook excelBook);
	/**
	 * 获取excel
	 */
	public ExcelBook getExcelBook(String excelId,int step);
}
