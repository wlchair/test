package com.acmr.excel.service;



import org.springframework.stereotype.Service;

import acmr.excel.ExcelException;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.util.ListHashMap;

import com.acmr.excel.model.Cell;
import com.acmr.excel.model.ColWidth;
import com.acmr.excel.model.RowHeight;
import com.acmr.excel.model.complete.GLXY;
import com.acmr.excel.model.complete.rows.ColOperate;
import com.acmr.excel.model.complete.rows.RowOperate;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.util.List;


/**
 * 单元格操作service
 * 
 * @author jinhr
 *
 */

@Service
public class CellService {
	/**
	 * 合并单元格
	 * 
	 * @param cell
	 *            Cell对象
	 */
	public void mergeCell(ExcelSheet excelSheet, Cell cell) {
		int firstRow = cell.getCoordinate().getStartSortY();
		int firstCol = cell.getCoordinate().getStartSortX();
		int lastRow = cell.getCoordinate().getEndSortY();
		int lastCol = cell.getCoordinate().getEndSortX();
		excelSheet.MergedRegions(firstRow, firstCol, lastRow, lastCol);
	}

	/**
	 * 单元格拆分
	 * 
	 * @param excelSheet
	 *            ExcelSheet对象
	 * @param cell
	 *            Cell对象
	 */

	public void splitCell(ExcelSheet excelSheet, Cell cell) {
		int firstRow = cell.getCoordinate().getStartSortY();
		int firstCol = cell.getCoordinate().getStartSortX();
		int lastRow = cell.getCoordinate().getEndSortY();
		int lastCol = cell.getCoordinate().getEndSortX();
		excelSheet.SplitRegions(firstRow, firstCol, lastRow, lastCol);
	}

	/**
	 * 增加行
	 * 
	 * @param sheet
	 *            SpreadSheet对象
	 * @param cell
	 *            Cell对象
	 */
	public void addRow(ExcelSheet excelSheet, RowOperate rowOperate) {
		excelSheet.insertRow(rowOperate.getRowSort());
	}

	/**
	 * 删除行
	 * 
	 * @param sheet
	 *            SpreadSheet对象
	 * @param cell
	 *            Cell对象
	 */
	public void deleteRow(ExcelSheet excelSheet, RowOperate rowOperate) {
		excelSheet.delRow(rowOperate.getRowSort());
	}

	/**
	 * 增加列
	 * 
	 * @param sheet
	 *            SpreadSheet对象
	 * @param cell
	 *            Cell对象
	 */
	public void addCol(ExcelSheet excelSheet, ColOperate colOperate) {
		excelSheet.insertColumn(colOperate.getColSort());
		List<ExcelColumn> colList = excelSheet.getCols();
		excelSheet.delColumn(colList.size()-1);
	}

	/**
	 * 删除列
	 * 
	 * @param sheet
	 *            SpreadSheet对象
	 * @param cell
	 *            Cell对象
	 */
	public void deleteCol(ExcelSheet excelSheet, ColOperate colOperate) {
		excelSheet.delColumn(colOperate.getColSort());
	}

	/**
	 * 调整列宽
	 * 
	 * @param excelSheet
	 *            SpreadSheet对象
	 * @param colAlais
	 *            列索引
	 * @param offset
	 *            偏移量
	 */
	public void controlColWidth(ExcelSheet excelSheet,ColWidth colWidth) {
		ListHashMap<ExcelColumn> colList = (ListHashMap<ExcelColumn>)excelSheet.getCols();
		int colIndex = colWidth.getColSort();
		int offsetWidth = colWidth.getOffset();
		colList.get(colIndex).setWidth(colList.get(colIndex).getWidth() + offsetWidth);
	}
	/**
	 * 调整列宽
	 * 
	 * @param excelSheet
	 *            SpreadSheet对象
	 * @param colAlais
	 *            列索引
	 * @param offset
	 *            偏移量
	 */
	public void colHide(ExcelSheet excelSheet,ColOperate colHide) {
		ListHashMap<ExcelColumn> colList = (ListHashMap<ExcelColumn>)excelSheet.getCols();
		int colIndex = colHide.getColSort();
		colList.get(colIndex).setColumnhidden(true);
	}
	/**
	 * 调整行高
	 * 
	 * @param excelSheet
	 *            SpreadSheet对象
	 * @param rowAlais
	 *            行索引
	 * @param offset
	 *            偏移量
	 */
	public void controlRowHeight(ExcelSheet excelSheet, RowHeight rowHeight) {
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>)excelSheet.getRows();
		int rowIndex = rowHeight.getRowSort();
		int offsetHeight = rowHeight.getOffset();
		rowList.get(rowIndex).setHeight(rowList.get(rowIndex).getHeight() + offsetHeight);
	}

}
