package com.acmr.excel.service;



import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.excel.pojo.ExcelSheetFreeze;
import acmr.util.ListHashMap;

import com.acmr.excel.model.Frozen;
import com.acmr.excel.model.complete.rows.ColOperate;
import com.acmr.excel.util.BinarySearch;

import java.util.List;
public class SheetService {
	/**
	 * 增加线，用于初始化时向下滚动
	 * 
	 * @param sheet
	 *            SheetElement
	 * @param rowNum
	 *            增加行数
	 */
	public void addRowLine(ExcelSheet sheet, int rowNum) {
		for (int i = 0; i < rowNum; i++) {
			sheet.addRow();
		}
	}

	/**
	 * 冻结
	 * 
	 * @param excelSheet
	 *            excelSheet对象
	 * @param frozenY
	 *            冻结横坐标
	 * @param frozenX
	 *            冻结纵坐标
	 * @param startY
	 *            开始点横坐标
	 * @param startX
	 *            开始点纵坐标
	 */
	public void frozen(ExcelSheet excelSheet,Frozen frozen) {
		ExcelSheetFreeze excelSheetFreeze = excelSheet.getFreeze();
		if (excelSheetFreeze == null) {
			excelSheetFreeze = new ExcelSheetFreeze();
			excelSheet.setFreeze(excelSheetFreeze);
		}
		int frozenYIndex = frozen.getFrozenSortY();
		int frozenXIndex = frozen.getFrozenSortX();
		excelSheetFreeze.setRow(frozenYIndex);
		excelSheetFreeze.setCol(frozenXIndex);
		excelSheetFreeze.setFirstrow(frozenYIndex);
		excelSheetFreeze.setFirstcol(frozenXIndex);
	}
	
	public void cancelColHide(ExcelSheet excelSheet,ColOperate colOperate) {
		List<ExcelColumn> colList = excelSheet.getCols();
		for(ExcelColumn excelColumn : colList){
			if(excelColumn.isColumnhidden()){
				excelColumn.setColumnhidden(false);;
			}
		}
	}
}
