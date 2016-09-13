package com.acmr.excel.service;



import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelCellStyle;
import acmr.excel.pojo.ExcelColor;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelFont;
import acmr.excel.pojo.ExcelFormat;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.excel.pojo.ExcelSheetFreeze;
import acmr.excel.pojo.Excelborder;
import acmr.util.ListHashMap;

import com.acmr.core.util.string.StringUtil;
import com.acmr.excel.convert.Convert;
import com.acmr.excel.dao.ExcelDao;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.OnlineExcel;
import com.acmr.excel.model.complete.Border;
import com.acmr.excel.model.complete.CompleteExcel;
import com.acmr.excel.model.complete.Content;
import com.acmr.excel.model.complete.CustomProp;
import com.acmr.excel.model.complete.Frozen;
import com.acmr.excel.model.complete.Glx;
import com.acmr.excel.model.complete.Gly;
import com.acmr.excel.model.complete.Occupy;
import com.acmr.excel.model.complete.OneCell;
import com.acmr.excel.model.complete.OperProp;
import com.acmr.excel.model.complete.ReturnParam;
import com.acmr.excel.model.complete.SheetElement;
import com.acmr.excel.model.complete.SpreadSheet;
import com.acmr.excel.model.complete.extend.HOneCell;
import com.acmr.excel.util.BinarySearch;
import com.acmr.excel.util.CellFormateUtil;
import com.acmr.excel.util.ExcelUtil;



/**
 * excel的还原操作service
 * 
 * @author jinhr
 *
 */
// @Service

public class ExcelService {
	private ExcelDao excelDao;

	public void setExcelDao(ExcelDao excelDao) {
		this.excelDao = excelDao;
	}

	/**
	 * 通过像素还原excel
	 * 
	 * @param rowBegin
	 *            开始行像素
	 * @param rowEnd
	 *            开始列像素
	 * @return CompleteExcel对象
	 */

	public SpreadSheet openExcel(SpreadSheet spreadSheet,
			ExcelSheet excelSheet, int rowBegin, int rowEnd,
			ReturnParam returnParam) {
		List<Gly> glyList = spreadSheet.getSheet().getGlY();
		List<Glx> glxList = spreadSheet.getSheet().getGlX();
		bookToOlExcelGlyList(excelSheet, glyList);
		bookToOlExcelGlxList(excelSheet, glxList);
		List<ExcelRow> rowList = excelSheet.getRows();
		int rowBeginIndex = BinarySearch.rowsBinarySearch(glyList, rowBegin);
		int rowEndIndex = BinarySearch.rowsBinarySearch(glyList, rowEnd);
		List<OneCell> newCellList = new ArrayList<OneCell>();
		spreadSheet.getSheet().setCells(newCellList);
		bookToOlExcelCellList(rowBeginIndex, rowEndIndex, rowList, glyList,
				glxList, excelSheet, newCellList);
		spreadSheet.getSheet().setGlY(
				glyList.subList(rowBeginIndex, rowEndIndex + 1));
		returnParam.setDataRowStartIndex(rowBeginIndex);
		return spreadSheet;
	}

	/**
	 * 通过别名加载excel
	 * 
	 * @return CompleteExcel对象
	 */

	public SpreadSheet openExcelByAlais(SpreadSheet spreadSheet,
			ExcelSheet excelSheet, String rowBeginAlais, String rowEndAlais,
			ReturnParam returnParam) {
		List<Gly> glyList = spreadSheet.getSheet().getGlY();
		List<Glx> glxList = spreadSheet.getSheet().getGlX();
		bookToOlExcelGlyList(excelSheet, glyList);
		bookToOlExcelGlxList(excelSheet, glxList);
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>) excelSheet
				.getRows();
		int rowBeginIndex = rowList.getMaps().get(rowBeginAlais);
		int rowEndIndex = rowList.getMaps().get(rowEndAlais);
		List<OneCell> newCellList = new ArrayList<OneCell>();
		spreadSheet.getSheet().setCells(newCellList);
		bookToOlExcelCellList(rowBeginIndex, rowEndIndex, rowList, glyList,
				glxList, excelSheet, newCellList);
		spreadSheet.getSheet().setGlY(
				glyList.subList(rowBeginIndex, rowEndIndex + 1));
		returnParam.setDataRowStartIndex(rowBeginIndex);
		return spreadSheet;
	}

	/**
	 * 通过Workbook转换为CompleteExcel
	 * 
	 * @param excel
	 *            CompleteExcel对象
	 * @return CompleteExcel对象
	 */

//	public CompleteExcel getExcel(Workbook workbook) {
//		Convert convert = new OnlineExcelConvert();
//		return convert.doConvertExcel(workbook);
//	}

	/**
	 * 保存excel
	 * 
	 * @param excel
	 *            OnlineExcel对象
	 */

	public void saveOrUpdateExcel(OnlineExcel excel) throws Exception {
		String excelId = excel.getExcelId();
		if(excelDao.getByExcelId(excelId) == 0){
			excelDao.saveExcel(excel);
		}
	}

	/**
	 * 获得所有的OnlineExcel对象
	 * 
	 * @return OnlineExcel对象集合
	 */

	public List<OnlineExcel> getAllExcel() {
		return excelDao.getAllExcel();
	}

	private void bookToOlExcelGlyList(ExcelSheet excelSheet, List<Gly> glyList) {
		List<ExcelRow> rowList = excelSheet.getRows();
		for (int i = 0; i < rowList.size(); i++) {
			Gly gly = new Gly();
			gly.setAliasY(rowList.get(i).getCode());
			// gly.setHeight(ExcelUtil.getPageHeight(rowList.get(i).getHeight()));
			gly.setHeight(rowList.get(i).getHeight());
			gly.setTop(getRowTop(glyList, i));
			gly.setIndex(i);
			glyList.add(gly);
		}

	}

	private void bookToOlExcelGlxList(ExcelSheet excelSheet, List<Glx> glxList) {
		List<ExcelColumn> colList = excelSheet.getCols();
		for (int i = 0; i < colList.size(); i++) {
			ExcelColumn excelColumn = colList.get(i);
			Glx glx = new Glx();
			glx.setAliasX(excelColumn.getCode());
			glx.setOriginWidth(excelColumn.getWidth());
			boolean hidden = excelColumn.isColumnhidden();
			if(hidden){
				glx.setHidden(true);
				glx.setWidth(0);
			}else{
				glx.setWidth(excelColumn.getWidth());
			}
			
			glx.setLeft(getColLeft(glxList, i));
			
			glx.setIndex(i);
			Map<String, String> colMap = excelColumn.getExps();
			//Content content = glxList.get(i).getOperProp().getContent();
			OperProp operProp = glx.getOperProp();
			Content content = operProp.getContent();
			String alignCol = colMap.get("align_vertical");
			content.setAlignCol(alignCol);
			String alignRow = colMap.get("align_level");
			content.setAlignRow(alignRow);
			String bd = colMap.get("font_weight");
			if (bd != null) {
				content.setBd(Boolean.valueOf(bd));
			} else {
				content.setBd(null);
			}
			String color = colMap.get("font_color");
			content.setColor(color);
			content.setFamily(colMap.get("font_family"));
			String italic = colMap.get("font_italic");
			if (italic != null) {
				content.setItalic(Boolean.valueOf(italic));
			} else {
				content.setItalic(null);
			}
			content.setSize(colMap.get("font_size"));
			content.setRgbColor(null);
			content.setTexts(null);
			content.setAlignLine(null);
//			CustomProp customProp = glxList.get(i).getOperProp()
//					.getCustomProp();
			CustomProp customProp = operProp.getCustomProp();
			customProp.setBackground(colMap.get("fill_bgcolor"));
			customProp.setFormat(colMap.get("format"));
			customProp.setCurrencySign(colMap.get("currency"));
			customProp.setDateFormat(colMap.get("dateFormat"));
			customProp.setIsValid(null);
			String decimalPoint = colMap.get("decimalPoint");
			if (decimalPoint != null) {
				customProp.setDecimal(Integer.valueOf(decimalPoint));
			} else {
				customProp.setDecimal(null);
			}
			String thousandPoint = colMap.get("thousandPoint");
			if (thousandPoint != null) {
				customProp.setThousands(Boolean.valueOf(thousandPoint));
			} else {
				customProp.setThousands(null);
			}

			customProp.setComment(colMap.get("comment"));
			//Border border = glxList.get(i).getOperProp().getBorder();
			Border border = operProp.getBorder();
			String bottom = colMap.get("bottom");
			String top = colMap.get("top");
			String left = colMap.get("left");
			String right = colMap.get("right");
			String all = colMap.get("all");
			String outer = colMap.get("outer");
			String none = colMap.get("none");
			if (bottom != null) {
				border.setBottom(Boolean.valueOf(bottom));
			} else {
				border.setBottom(null);
			}
			if (top != null) {
				border.setTop(Boolean.valueOf(top));
			} else {
				border.setTop(null);
			}
			if (left != null) {
				border.setLeft(Boolean.valueOf(left));
			} else {
				border.setLeft(null);
			}
			if (right != null) {
				border.setRight(Boolean.valueOf(right));
			} else {
				border.setRight(null);
			}
			if (all != null) {
				border.setAll(Boolean.valueOf(all));
			} else {
				border.setAll(null);
			}
			if (outer != null) {
				border.setOuter(Boolean.valueOf(outer));
			} else {
				border.setOuter(null);
			}
			if (none != null) {
				border.setNone(Boolean.valueOf(none));
			} else {
				border.setNone(null);
			}
			
			operProp.setContent(content);
			operProp.setBorder(border);
			operProp.setContent(content);
			glx.setOperProp(operProp);
			glxList.add(glx);
		}
	}

	private int getColLeft(List<Glx> glxList, int i) {
		if (i == 0) {
			return 0;
		}
		Glx glx = glxList.get(i - 1);
		int tempWidth = glx.getWidth();
		if(glx.getWidth() == 0){
			tempWidth = -1;
		}
		return glx.getLeft() + tempWidth + 1;
	}

	private int getRowTop(List<Gly> glyList, int i) {
		if (i == 0) {
			return 0;
		}
		Gly gly = glyList.get(i - 1);
		return gly.getTop() + gly.getHeight() + 1;
	}

	private void bookToOlExcelCellList(int rowBeginIndex, int rowEndIndex,
			List<ExcelRow> rowList, List<Gly> glyList, List<Glx> glxList,
			ExcelSheet excelSheet, List<OneCell> newCellList) {
		getCellList(rowBeginIndex, rowEndIndex, rowList, glyList, glxList, excelSheet, newCellList);
	}

	/**
	 * 获得普通单元格
	 * 
	 * @param rowBeginIndex
	 * @param rowEndIndex
	 * @param rowList
	 * @param glyList
	 * @param glxList
	 * @param excelSheet
	 * @param newCellList
	 */

	private void getCellList(int rowBeginIndex, int rowEndIndex,List<ExcelRow> rowList, List<Gly> glyList, List<Glx> glxList,
			ExcelSheet excelSheet, List<OneCell> newCellList) {
		for (int i = rowBeginIndex; i <= rowEndIndex; i++) {
			ExcelRow excelRow = rowList.get(i);
			if (excelRow != null) {
				List<ExcelCell> cellList = excelRow.getCells();
				for (int j = 0; j < cellList.size(); j++) {
					ExcelCell excelCell = cellList.get(j);
					if (excelCell != null) {
						OneCell cell = new OneCell();
						String highLight = excelCell.getExps().get(Constant.HIGHLIGHT);
						if ("true".equals(highLight)) {
							cell.setHighlight(true);
						} else {
							cell.setHighlight(false);
						}
						setCellProperty(excelCell, cell, i, j, glyList,glxList, excelSheet);
						if (excelSheet.getCols().get(j).isColumnhidden()) {
							int colspan = excelCell.getColspan();
							//int rowspan = excelCell.getRowspan();
							if (colspan > 1) {
								int[] mfCell = excelSheet.getMergFirstCell(i, j);
								int temp = mfCell[1];
								boolean flag = true;
								for (int k = temp; k < temp + colspan; k++) {
									if (!excelSheet.getCols().get(k).isColumnhidden()) {
										flag = false;
									}
								}
								if (flag) {
									cell.setHidden(true);
								}
							}else{
								cell.setHidden(true);
							}
						}
						newCellList.add(cell);
					}
				}
				Map<String, String> colMap = excelRow.getExps();
				Content content = glyList.get(i).getOperProp().getContent();
				String alignCol = colMap.get("align_vertical");
				content.setAlignCol(alignCol);
				String alignRow = colMap.get("align_level");
				content.setAlignRow(alignRow);
				String bd = colMap.get("font_weight");
				if (bd != null) {
					content.setBd(Boolean.valueOf(bd));
				}else{
					content.setBd(null);
				}
				String color = colMap.get("font_color");
				content.setColor(color);
				content.setFamily(colMap.get("font_family"));
				String italic = colMap.get("font_italic");
				if (italic != null) {
					content.setItalic(Boolean.valueOf(italic));
				} else {
					content.setItalic(null);
				}
				content.setSize(colMap.get("font_size"));
				content.setRgbColor(null);
				content.setTexts(null);
				CustomProp customProp = glyList.get(i).getOperProp().getCustomProp();
				customProp.setBackground(colMap.get("fill_bgcolor"));
				customProp.setFormat(colMap.get("format"));
				customProp.setCurrencySign(colMap.get("currency"));
				customProp.setDateFormat(colMap.get("dateFormat"));
				String decimalPoint = colMap.get("decimalPoint");
				if (decimalPoint != null) {
					customProp.setDecimal(Integer.valueOf(decimalPoint));
				} else {
					customProp.setDecimal(null);
				}
				String thousandPoint = colMap.get("thousandPoint");
				if (thousandPoint != null) {
					customProp.setThousands(Boolean.valueOf(thousandPoint));
				} else {
					customProp.setThousands(null);
				}
				
				customProp.setComment(colMap.get("comment"));
				Border border = glyList.get(i).getOperProp().getBorder();
				String bottom = colMap.get("bottom");
				String top = colMap.get("top");
				String left = colMap.get("left");
				String right = colMap.get("right");
				String all = colMap.get("all");
				String outer = colMap.get("outer");
				String none = colMap.get("none");
				if(bottom != null){
					border.setBottom(Boolean.valueOf(bottom));
				}else{
					border.setBottom(null);
				}
				if(top != null){
					border.setTop(Boolean.valueOf(top));
				}else{
					border.setTop(null);
				}
				if(left != null){
					border.setLeft(Boolean.valueOf(left));
				}else{
					border.setLeft(null);
				}
				if(right != null){
					border.setRight(Boolean.valueOf(right));
				}else{
					border.setRight(null);
				}
				if(all != null){
					border.setAll(Boolean.valueOf(all));
				}else{
					border.setAll(null);
				}
				if(outer != null){
					border.setOuter(Boolean.valueOf(outer));
				}else{
					border.setOuter(null);
				}
				if(none != null){
					border.setNone(Boolean.valueOf(none));
				}else{
					border.setNone(null);
				}
			}
		}
	}


	/**
	 * 设置单元格常规属性
	 * 
	 * @param excelCell
	 * @param cell
	 * @param i
	 * @param j
	 * @param glyList
	 * @param glxList
	 * @param excelSheet
	 */

	private void setCellProperty(ExcelCell excelCell, OneCell cell, int i,int j, List<Gly> glyList, List<Glx> glxList, ExcelSheet excelSheet) {
		ExcelCellStyle excelCellStyle = excelCell.getCellstyle();
		Border border = cell.getBorder();
		Excelborder excelTopborder = excelCellStyle.getTopborder();
		if(excelTopborder != null){
			short topBorder = excelTopborder.getSort();
			if (topBorder > 0) {
				border.setTop(true);
			}
		}
		Excelborder excelBottomborder  = excelCellStyle.getBottomborder();
		if(excelBottomborder != null){
			short bottomBorder = excelBottomborder.getSort();
			if (bottomBorder >0) {
				border.setBottom(true);
			}
		}
		Excelborder excelLeftborder = excelCellStyle.getLeftborder();
		if(excelLeftborder != null){
			short leftBorder = excelLeftborder.getSort();
			if (leftBorder > 0) {
				border.setLeft(true);
			}
		}
		Excelborder excelRightborder = excelCellStyle.getRightborder();
		if(excelRightborder != null){
			short rightBorder = excelRightborder.getSort();
			if (rightBorder > 0) {
				border.setRight(true);
			}
		}
		cell.setWordWrap(excelCellStyle.isWraptext());
		Content content = cell.getContent();
		switch (excelCellStyle.getAlign()) {
		case 1:
			content.setAlignRow("left");
			break;
		case 2:
			content.setAlignRow("center");
			break;
		case 3:
			content.setAlignRow("right");
			break;

		default:
			//content.setAlignRow("left");
			break;
		}
		switch (excelCellStyle.getValign()) {
		case 0:
			content.setAlignCol("top");
			break;
		case 1:
			content.setAlignCol("middle");
			break;
		case 2:
			content.setAlignCol("bottom");
			break;

		default:
			content.setAlignCol("middle");
			break;
		}
		ExcelFont excelFont = excelCellStyle.getFont();
		content.setFamily(excelFont.getFontname());
		if (700 == excelFont.getBoldweight()) {
			content.setBd(true);
		} else {
			content.setBd(false);
		}
		content.setItalic(excelFont.isItalic());
//		ExcelColor fontColor = excelFont.getColor();
//		if (fontColor != null) {
//			int[] rgb = fontColor.getRGBInt();
//			content.setColor(ExcelUtil.getRGB(rgb));
//			content.setColor("rgb(0,0,0)");
//		}
		content.setSize(excelFont.getSize() / 20 + "");
		content.setTexts(excelCell.getText());
		CustomProp customProp = cell.getCustomProp();
		//if ("true".equals(excelSheet.getExps().get("ifUpload"))) {
		if (excelCell.getExps().get("format") == null) {
		//if (("".equals(excelCell.getShowText())|| excelCell.getShowText() == null) && !"".equals(excelCell.getText())) {
			CellFormateUtil.setShowText(excelCell, content,customProp);
			String displayText = ExcelFormat.getShowText(excelCell);
			content.setDisplayTexts(displayText);
		} else {
			String format = excelCell.getExps().get("format");
			customProp.setFormat(format);
			customProp.setThousands(Boolean.valueOf(excelCell.getExps().get("thousandPoint")));
			if ("date".equals(format)) {
				customProp.setDateFormat(excelCell.getExps().get("dataFormate"));
			}
			String decimalPoint = excelCell.getExps().get("decimalPoint");
			if (!StringUtil.isEmpty(decimalPoint)) {
				customProp.setDecimal(Integer.valueOf(decimalPoint));
			}
			customProp.setCurrencySign(excelCell.getExps().get("currencySymbol"));
		}
		String showText = excelCell.getShowText();
		if(showText == null){
			content.setDisplayTexts("");
		}else{
			content.setDisplayTexts(showText);
		}
		if(!StringUtil.isEmpty(excelCell.getMemo())){
			customProp.setComment(excelCell.getMemo());
		}
		
		
		ExcelColor fontColor = excelFont.getColor();
		if (fontColor != null) {
			int[] rgb = fontColor.getRGBInt();
			content.setColor(ExcelUtil.getRGB(rgb));
			//content.setColor("rgb(0,0,0)");
		}
		// ExcelColor bgColor = excelCellStyle.getBgcolor();
		ExcelColor fgColor = excelCellStyle.getFgcolor();
		// if(bgColor != null){
		// int[] rgb = bgColor.getRGBInt();
		// customProp.setBackground(ExcelUtil.getRGB(rgb));
		// }else if(bgColor == null && fgColor != null){
		if (fgColor != null) {
			int[] rgb = fgColor.getRGBInt();
			customProp.setBackground(ExcelUtil.getRGB(rgb));
		}
		Occupy occupy = cell.getOccupy();
		int rowspan = excelCell.getRowspan();
		int colspan = excelCell.getColspan();
		int[] firstMergeCell = excelSheet.getMergFirstCell(i, j);
		if (firstMergeCell == null) {
			occupy.getX().add(glxList.get(j).getAliasX());
			occupy.getY().add(glyList.get(i).getAliasY());
		} else {
			int firstRow = firstMergeCell[0];
			int firstCol = firstMergeCell[1];
			for (int m = firstRow; m < firstRow + rowspan; m++) {
				occupy.getY().add(glyList.get(m).getAliasY());
			}
			for (int m = firstCol; m < firstCol + colspan; m++) {
				occupy.getX().add(glxList.get(m).getAliasX());
			}
		}
		
		
		
		
		
		
		
		
		
		
	}

	/**
	 * 非冻结定位还原excel
	 * 
	 * @param height
	 *            高度
	 * @param returnParam
	 *            返回参数
	 * @return
	 */

	public SpreadSheet positionExcel(ExcelSheet excelSheet, SpreadSheet spreadSheet, int height, ReturnParam returnParam) {
		spreadSheet.setName(excelSheet.getName());

		List<Gly> glyList = spreadSheet.getSheet().getGlY();
		List<Glx> glxList = spreadSheet.getSheet().getGlX();
		bookToOlExcelGlyList(excelSheet, glyList);
		bookToOlExcelGlxList(excelSheet, glxList);
		List<ExcelRow> rowList = excelSheet.getRows();
		// List<ExcelColumn> colList = excelSheet.getCols();
		int minTop = glyList.get(0).getTop();
		Gly glyTop = glyList.get(glyList.size() - 1);
		int maxTop = glyTop.getTop() + glyTop.getHeight();
		returnParam.setMaxPixel(maxTop);
		int startAlaisPixel = 0;
		int Offset = startAlaisPixel - minTop;
		int startPixel = Offset < 200 ? Offset : startAlaisPixel - 200;
		int endPixel = 0;
		if (maxTop < height) {
			endPixel = maxTop;
		} else {
			endPixel = startPixel + height + 200;
		}
		// int oldrowBeginIndex = BinarySearch.rowsBinarySearch(startPixel,
		// glyList, 0, glyList.size()-1);
		int rowBeginIndex = BinarySearch.rowsBinarySearch(glyList, startPixel);
		// int oldrowEndIndex = BinarySearch.rowsBinarySearch(endPixel, glyList,
		// 0, glyList.size()-1);
		int rowEndIndex = BinarySearch.rowsBinarySearch(glyList, endPixel);
		List<OneCell> newCellList = new ArrayList<OneCell>();
		spreadSheet.getSheet().setCells(newCellList);
		bookToOlExcelCellList(rowBeginIndex, rowEndIndex, rowList, glyList,glxList, excelSheet, newCellList);
		spreadSheet.getSheet().setGlY(glyList.subList(rowBeginIndex, rowEndIndex + 1));
		spreadSheet.getSheet().setGlX(glxList);
		returnParam.setDataRowStartIndex(rowBeginIndex);
		returnParam.setDisplayRowStartAlias("1");
		returnParam.setDisplayColStartAlias("1");
		ExcelSheetFreeze excelfrozen = excelSheet.getFreeze();
		if (excelfrozen != null) {
			Frozen frozen = spreadSheet.getSheet().getFrozen();
			frozen.setRowIndex(excelfrozen.getRow() + 1 + "");
			frozen.setColIndex(excelfrozen.getCol() + 1 + "");
			frozen.setState("1");
		}
		return spreadSheet;
	}


	/**
	 * 通过id获得OnlineExcel对象
	 * 
	 * @param id
	 * @return OnlineExcel
	 */

	public String getExcel(String excelId) {
		try {
			OnlineExcel oe = excelDao.getJsonObjectByExcelId(excelId);
			return oe.getExcelObject();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}


	/**
	 * 改变宽度或高度
	 * 
	 * @param excelBook
	 */

	public void changeHeightOrWidth(ExcelBook excelBook) {
		ExcelSheet excelSheet = excelBook.getSheets().get(0);
		List<ExcelColumn> colList = excelSheet.getCols();
		List<ExcelRow> rowList = excelSheet.getRows();
		for (ExcelColumn excelColumn : colList) {
			excelColumn.setWidth(excelColumn.getWidth() * 40);
		}
		for (ExcelRow excelRow : rowList) {
			excelRow.setHeight(excelRow.getHeight() * 18);
		}
	}
}
