package com.acmr.excel.service;



import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.acmr.core.util.string.StringUtil;
import com.acmr.excel.model.complete.BaseCell;
import com.acmr.excel.model.complete.Border;
import com.acmr.excel.model.complete.CompleteExcel;
import com.acmr.excel.model.complete.Content;
import com.acmr.excel.model.complete.CustomProp;
import com.acmr.excel.model.complete.Frozen;
import com.acmr.excel.model.complete.Glx;
import com.acmr.excel.model.complete.Gly;
import com.acmr.excel.model.complete.Occupy;
import com.acmr.excel.model.complete.OneCell;
import com.acmr.excel.model.complete.Position;
import com.acmr.excel.model.complete.SpreadSheet;

import java.awt.Color;
import java.util.List;
import java.util.Map;

public class DownLoadExcelService {
	/**
	 * excel下载
	 * 
	 * @param excel
	 *            excel对象
	 * @return
	 */
	public Workbook downloadExcel(CompleteExcel excel) {
		Workbook wb = new XSSFWorkbook();
		List<SpreadSheet> spreadSheets = excel.getSpreadSheet();
		if (spreadSheets == null || spreadSheets.size() == 0) {
			return wb;
		}
		for (SpreadSheet spreadSheet : spreadSheets) {
			if (spreadSheet == null) {
				continue;
			}
			Sheet sheet = wb.createSheet(spreadSheet.getName());

			// Map<String, List<Map<String, Integer>>> aliasX =
			// spreadSheet.getSheet().getPosi().getStrandX().getAliasX();
			Position position = spreadSheet.getSheet().getPosi();
			List<Glx> glx = spreadSheet.getSheet().getGlX();
			List<Gly> gly = spreadSheet.getSheet().getGlY();
			Map<String, Map<String, Integer>> aliasY = position.getStrandY()
					.getAliasY();
			List<OneCell> oneCells = spreadSheet.getSheet().getCells();
			// if(oneCells == null || oneCells.size() == 0){
			// return wb;
			// }
			int i = 0;
			for (Gly y : gly) {
				String aliasy = y.getAliasY();
				Row row = sheet.createRow(i);
				int height = y.getHeight();
				row.setHeight((short) (height * 18));// 20
				i++;
				int j = 0;
				for (Glx x : glx) {
					String aliasx = x.getAliasX();
					Cell cell = row.createCell(j);
					int weight = x.getWidth();
					sheet.setColumnWidth(j, weight * 40);// 256
					Map<String, Integer> xMap = aliasY.get(aliasy);
					if (xMap != null) {
						Integer index = xMap.get(aliasx);
						if (index != null) {
							OneCell oneCell = oneCells.get(index);
							// if(oneCell.isUsed() == false){
							// this.setMergedRegion(oneCell, sheet, glx, gly);
							// this.setCellAttr(wb, cell, oneCell);
							// oneCell.setUsed(true);
							// }
						}
					}
					j++;
				}
			}
			Frozen frozen = spreadSheet.getSheet().getFrozen();
			if ("1".equals(frozen.getState())) {
				this.setFreezePane(frozen, sheet, gly, glx);
			}
		}
		return wb;
	}

	// private void removeRepeatPosition(OneCell oneCell) {
	// List<String> ys = oneCell.getOccupy().getY();
	// List<String> xs = oneCell.getOccupy().getX();
	// if (ys.size() > 1 || xs.size() > 1) {
	// for (int k = 0; k < ys.size(); k++) {
	// for (int l = 0; l < xs.size(); l++) {
	// if (k == 0 && l == 0) {
	// continue;
	// }
	// oneCell.setUsed(true);
	// }
	// }
	// }
	// }
	/**
	 * 设置单元格各种样式
	 * 
	 * @param wb
	 * @param cell
	 * @param oc
	 */

	private void setCellAttr(Workbook wb, Cell cell, BaseCell oc) {
		if (cell == null || wb == null || oc == null) {
			return;
		}
		XSSFCellStyle cellStyle = (XSSFCellStyle) wb.createCellStyle();
		XSSFFont font = (XSSFFont) wb.createFont();

		Content content = oc.getContent();
		Border border = oc.getBorder();
		CustomProp customProp = oc.getCustomProp();
		font.setFontName(content.getFamily());
		cell.setCellValue(content.getTexts());
		String size = content.getSize();
		size = size.substring(0, size.length() - 2);
		int pointIndex = size.indexOf(".");
		if (pointIndex != -1) {
			size = size.substring(0, pointIndex);
		}
		font.setFontHeightInPoints(Short.valueOf(size));
		String color = content.getColor();
		if (!StringUtil.isEmpty(color) && !"rgb(0,0,0)".equals(color)) {
			int r = getRGB(color, 0);
			int g = getRGB(color, 1);
			int b = getRGB(color, 2);
			font.setColor(new XSSFColor(new Color(r, g, b))); // 字体颜色
		}
		// font.setColor(IndexedColors.BLACK.index);
		font.setItalic(content.getItalic()); // 是否为斜体
//		if (content.isBd()) {
//			font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);// 粗体显示
//		}

		// cellStyle.setBorderLeft(paramShort);

		cellStyle.setFont(font);// 选择需要用到的字体格式
		// 边框样式
		if (!border.getNone()) {
			if (border.getBottom()) {
				cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);// 下边框
			}
			if (border.getLeft()) {
				cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);// 左边框
			}
			if (border.getTop()) {
				cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);// 上边框
			}
			if (border.getRight()) {
				cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);// 右边框
			}
			if (border.getAll()) {
				cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);// 下边框
				cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);// 左边框
				cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);// 上边框
				cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);// 右边框
			}
		}
		String bgRgb = customProp.getBackground();
		if (!StringUtil.isEmpty(bgRgb) && !"#fff".equals(bgRgb)) {
			int r = getRGB(bgRgb, 0);
			int g = getRGB(bgRgb, 1);
			int b = getRGB(bgRgb, 2);
			cellStyle.setFillForegroundColor(new XSSFColor(new Color(r, g, b)));
			cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		}
		String vertical = content.getAlignRow();
		String align = content.getAlignLine();
		cellStyle.setAlignment(getAlign(align));
		cellStyle.setVerticalAlignment(getVertical(vertical));
		cell.setCellStyle(cellStyle);

	}

	private int getRGB(String color, int index) {
		color = color.substring(4, color.length() - 1);
		String[] rgb = color.split(",");
		return Integer.valueOf(rgb[index].trim());
	}

	/**
	 * 获得单元格水平位置
	 * 
	 * @param align
	 * @return
	 */

	private Short getAlign(String align) {
		Short newAlign = CellStyle.ALIGN_LEFT;
		switch (align) {
		case "left":
			newAlign = CellStyle.ALIGN_LEFT;
			break;
		case "center":
			newAlign = CellStyle.ALIGN_CENTER;
			break;
		case "middle":
			newAlign = CellStyle.ALIGN_CENTER;
			break;
		case "right":
			newAlign = CellStyle.ALIGN_RIGHT;
			break;
		default:
			break;
		}
		return newAlign;
	}

	/**
	 * 获得单元格垂直位置
	 * 
	 * @param align
	 * @return
	 */

	private Short getVertical(String align) {
		Short newAlign = CellStyle.VERTICAL_CENTER;
		switch (align) {
		case "top":
			newAlign = CellStyle.VERTICAL_TOP;
			break;
		case "middle":
			newAlign = CellStyle.VERTICAL_CENTER;
			break;
		case "bottom":
			newAlign = CellStyle.VERTICAL_BOTTOM;
			break;
		default:
			break;
		}
		return newAlign;
	}

	//
	// }

	private void setMergedRegion(OneCell oneCell, Sheet sheet, List<Glx> glx,
			List<Gly> gly) {
		Occupy occupy = oneCell.getOccupy();
		List<String> xs = occupy.getX();
		List<String> ys = occupy.getY();
		if (xs.size() > 1 || ys.size() > 1) {
			int firstCol = getXIndex(xs.get(0), glx);
			int lastCol = getXIndex(xs.get(xs.size() - 1), glx);
			int firstRow = getYIndex(ys.get(0), gly);
			int lastRow = getYIndex(ys.get(ys.size() - 1), gly);
			sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow,
					firstCol, lastCol));
		}

	}

	private int getXIndex(String alias, List<Glx> glx) {
		int index = -1;
		for (int i = 0; i < glx.size(); i++) {
			if (alias.equals(glx.get(i).getAliasX())) {
				index = i;
			}
		}
		return index;
	}

	private int getYIndex(String alias, List<Gly> gly) {
		int index = -1;
		for (int i = 0; i < gly.size(); i++) {
			if (alias.equals(gly.get(i).getAliasY())) {
				index = i;
			}
		}
		return index;
	}

	/**
	 * 冻结
	 */

	private void setFreezePane(Frozen frozen, Sheet sheet, List<Gly> glys,
			List<Glx> glxs) {
		if (frozen == null || sheet == null) {
			return;
		}
		String yAlais = frozen.getRowIndex();
		String xAlais = frozen.getColIndex();
		int rVal = getYIndex(yAlais, glys);
		int cVal = getXIndex(xAlais, glxs);
		sheet.createFreezePane(cVal, rVal);
	}

	// public static String getColor(String color){
	// Class clazz = XSSFColor.class;
	// Class[] clazzs = clazz.getDeclaredClasses();
	// for(Class ){
	//
	// }
	// return "";
	// }

	// public static void main(String[] args) {
	// String str = "#F0FFF0";
	// //处理把它转换成十六进制并放入一个数
	// int[] color=new int[3];
	// color[0]=Integer.parseInt(str.substring(1, 3), 16);
	// color[1]=Integer.parseInt(str.substring(3, 5), 16);
	// color[2]=Integer.parseInt(str.substring(5, 7), 16);
	// System.out.println(color);
	// }

}
