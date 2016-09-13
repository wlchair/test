package com.acmr.excel.service;



import acmr.excel.pojo.Constants.CELLTYPE;
import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelColor;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.util.ListHashMap;

import com.acmr.excel.model.OuterPasteData;
import com.acmr.excel.model.Paste;
import com.acmr.excel.model.copy.Copy;
import java.io.UnsupportedEncodingException;
import java.util.List;
import java.util.Map;
public class PasteService {
	/**
	 * 是否可以复制
	 * 
	 * @param copy
	 * @param excelBook
	 * @return
	 */
	public boolean canCopy(Copy copy, ExcelBook excelBook) {
		boolean canPaste = true;
		ExcelSheet excelSheet = excelBook.getSheets().get(0);
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>) excelSheet.getRows();
		int startRowIndex = copy.getOrignal().getStartRowSort();
		int endRowIndex = copy.getOrignal().getEndRowSort();
		int startColIndex = copy.getOrignal().getStartColSort();
		int endColIndex = copy.getOrignal().getEndColSort();
		int targetRowIndex = copy.getTarget().getRowSort();
		int targetColIndex = copy.getTarget().getColSort();
		int rowRange = endRowIndex - startRowIndex;
		int colRange = endColIndex - startColIndex;
		for (int i = targetRowIndex; i < targetRowIndex + rowRange; i++) {
			ExcelRow excelRow = rowList.get(i);
			for (int j = targetColIndex; j < targetColIndex + colRange; j++) {
				ExcelCell excelCell = excelRow.getCells().get(j);
				if (excelCell == null) {
					continue;
				}
				int colspan = excelCell.getColspan();
				int rowspan = excelCell.getRowspan();
				if (colspan != 1 || rowspan != 1) {
					canPaste = false;
					break;
				}
			}
		}
		return canPaste;
	}

	/**
	 * 粘贴
	 * 
	 * @param paste
	 *            paste对象
	 * @param excelBook
	 *            excelBook对象
	 */

	public void data(Paste paste, ExcelBook excelBook) {
		ExcelSheet excelSheet = excelBook.getSheets().get(0);
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>) excelSheet.getRows();
		ListHashMap<ExcelColumn> colList = (ListHashMap<ExcelColumn>) excelSheet.getCols();
		Map<String, Integer> rowMap = rowList.getMaps();
		Map<String, Integer> colMap = colList.getMaps();
		List<OuterPasteData> pasteList = paste.getPasteData();
		for (OuterPasteData outerPasteData : pasteList) {
			int rowIndex = outerPasteData.getRowSort();
			ExcelRow excelRow = rowList.get(rowIndex);
			int colIndex = outerPasteData.getColSort();
			List<ExcelCell> cellList = excelRow.getCells();
			ExcelCell excelCell = cellList.get(colIndex);
			if (excelCell == null) {
				excelCell = new ExcelCell();
				excelCell.getCellstyle().setBgcolor(new ExcelColor(255, 255, 255));
				cellList.set(colIndex, excelCell);
			}
			String text = outerPasteData.getText();
				//text = java.net.URLDecoder.decode(text, "utf-8");
				excelCell.setText(text);
				excelCell.setType(CELLTYPE.STRING);
				excelCell.setValue(text);
		}
	}

	/**
	 * 复制粘贴
	 * 
	 * @param copy
	 * @param excelBook
	 */
	public void copy(Copy copy, ExcelBook excelBook) {
		copyOrCut(copy, excelBook, null);
	}

	/**
	 * 剪切粘贴
	 * 
	 * @param copy
	 * @param excelBook
	 */

	public void cut(Copy copy, ExcelBook excelBook) {
		copyOrCut(copy, excelBook, "cut");
	}

	/**
	 * 复制或剪切
	 * 
	 * @param copy
	 * @param excelBook
	 * @param flag
	 *            copy:复制 cut:剪切
	 */

	private void copyOrCut(Copy copy, ExcelBook excelBook, String flag) {
		ExcelSheet excelSheet = excelBook.getSheets().get(0);
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>) excelSheet.getRows();
		int startRowIndex = copy.getOrignal().getStartRowSort();
		int endRowIndex = copy.getOrignal().getEndRowSort();
		int startColIndex = copy.getOrignal().getStartColSort();
		int endColIndex = copy.getOrignal().getEndColSort();
		int targetRowIndex = copy.getTarget().getRowSort();
		int targetColIndex = copy.getTarget().getColSort();
		for (int i = startRowIndex; i <= endRowIndex; i++) {
			ExcelRow excelRow = rowList.get(i);
			int tempColIndex = targetColIndex;
			for (int j = startColIndex; j <= endColIndex; j++) {
				ExcelCell excelCell = excelRow.getCells().get(j);
				if(excelCell == null){
					excelCell = new ExcelCell();
				}
				ExcelCell newExcelCell = excelCell.clone();
				rowList.get(targetRowIndex).set(tempColIndex, newExcelCell);
				if ("cut".equals(flag)) {
					rowList.get(i).set(j, null);
				}
				tempColIndex++;
			}
			targetRowIndex++;
		}
	}

	/**
	 * 是否可以粘贴
	 * 
	 * @param isAblePaste
	 * @param excelBook
	 * @return
	 */

	public boolean isAblePaste(Paste paste, ExcelBook excelBook) {
		boolean canPaste = true;
		ExcelSheet excelSheet = excelBook.getSheets().get(0);
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>) excelSheet.getRows();
		int startRowIndex = paste.getStartRowSort();
		int startColIndex = paste.getStartColSort();
		int rowRange = paste.getRowLen();
		int colRange = paste.getColLen();
		for (int i = startRowIndex; i < startRowIndex + rowRange; i++) {
			ExcelRow excelRow = rowList.get(i);
			for (int j = startColIndex; j < startColIndex + colRange; j++) {
				ExcelCell excelCell = excelRow.getCells().get(j);
				if (excelCell == null) {
					continue;
				}
				int colspan = excelCell.getColspan();
				int rowspan = excelCell.getRowspan();
				if (colspan != 1 || rowspan != 1) {
					canPaste = false;
					break;
				}
			}
		}
		return canPaste;
	}
	
	public boolean isCopyPaste(Copy copy, ExcelBook excelBook){
		boolean canPaste = true;
		ExcelSheet excelSheet = excelBook.getSheets().get(0);
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>) excelSheet.getRows();
		int startRowIndex = copy.getOrignal().getStartRowSort();
		int endRowIndex = copy.getOrignal().getEndRowSort();
		int startColIndex = copy.getOrignal().getStartColSort();
		int endColIndex = copy.getOrignal().getEndColSort();
		int rowRange = endRowIndex - startRowIndex;
		int colRange = endColIndex - startColIndex;
		int targetStartRowIndex = copy.getTarget().getRowSort();
		int targetStartColIndex = copy.getTarget().getColSort();
		for (int i = targetStartRowIndex; i <= targetStartRowIndex + rowRange; i++) {
			ExcelRow excelRow = rowList.get(i);
			for (int j = targetStartColIndex; j <= targetStartColIndex + colRange; j++) {
				ExcelCell excelCell = excelRow.getCells().get(j);
				if (excelCell == null) {
					continue;
				}
				int colspan = excelCell.getColspan();
				int rowspan = excelCell.getRowspan();
				if (colspan != 1 || rowspan != 1) {
					canPaste = false;
					break;
				}
			}
		}
		return canPaste;
	}
	
	

}
