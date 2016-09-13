package com.acmr.excel.action;

import java.io.IOException;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import net.spy.memcached.MemcachedClient;

import org.springframework.web.servlet.ModelAndView;

import acmr.excel.pojo.ExcelBook;

import com.acmr.excel.action.excelbase.ExcelBaseAction;
import com.acmr.excel.model.Cell;
import com.acmr.excel.model.ColWidth;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.RowHeight;
import com.acmr.excel.model.complete.rows.ColOperate;
import com.acmr.excel.model.complete.rows.RowOperate;
import com.acmr.excel.service.HandleExcelService;
import com.acmr.excel.util.MemcacheUtil;
import com.acmr.helper.model.JSONReturnData;

/**
 * 单元格操作
 * 
 * @author jinhr
 *
 */
// @Controller
// @RequestMapping("/cell")
public class CellAction extends ExcelBaseAction {
	private HandleExcelService handleExcelService;
	private MemcachedClient memcachedClient;

	public void setHandleExcelService(HandleExcelService handleExcelService) {
		this.handleExcelService = handleExcelService;
	}


	public void setMemcachedClient(MemcachedClient memcachedClient) {
		this.memcachedClient = memcachedClient;
	}

	/**
	 * 创建单元格
	 * 
	 * @throws IOException
	 */
	// @RequestMapping("/create")
	public void create(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
//		Cell cell = getJsonDataParameter(req, Cell.class);
//		JSONReturnData data = new JSONReturnData("");
//		String excelId = cell.getExcelId();
//		ExcelBook excel = (ExcelBook) MemcacheUtil.get(excelId, memcachedClient);
//		if (excel != null) {
//			String row = cell.getCoordinate().getStartY();
//			String col = cell.getCoordinate().getStartX();
//			int sheetId = 0;
//			handleExcelService.createCell(excel, sheetId, row, col);
//			MemcacheUtil.set(excelId, memcachedClient, excel);
//			data.setReturncode(Constant.SUCCESS_CODE);
//		} else {
//			data.setReturncode(Constant.CACHE_INVALID_CODE);
//			data.setReturndata(Constant.CACHE_INVALID_MSG);
//		}
		//this.sendJson(resp, data);
	}

	/**
	 * 合并单元格
	 * 
	 * @throws IOException
	 */

	public void merge(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp,cell,OperatorConstant.merge);
	}

	/**
	 * 单元格拆分
	 * 
	 * @throws IOException
	 */
	public void merge_delete(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp,cell,OperatorConstant.mergedelete);
	}

	/**
	 * 边框操作
	 * 
	 * @throws IOException
	 */

	public void frame(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp,cell,OperatorConstant.frame);
	}


	/**
	 * 水平对齐
	 * 
	 * @throws IOException
	 */

	public void align_level(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		// 接收参数，定义返回
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp,cell,OperatorConstant.alignlevel);
	}

	/**
	 * 垂直对齐
	 * 
	 * @throws IOException
	 */
	public void align_vertical(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		// 接收参数，定义返回
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp,cell,OperatorConstant.alignvertical);
	}

	/**
	 * 插入行
	 * 
	 * @throws IOException
	 */

	public void rows_insert(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		RowOperate rowOperate = getJsonDataParameter(req, RowOperate.class);
		this.assembleData(req, resp,rowOperate,OperatorConstant.rowsinsert);
	}

	/**
	 * 删除行
	 * 
	 * @throws IOException
	 */

	public void rows_delete(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		RowOperate rowOperate = getJsonDataParameter(req, RowOperate.class);
		this.assembleData(req, resp,rowOperate,OperatorConstant.rowsdelete);
	}

	/**
	 * 列操作
	 * 
	 * @throws IOException
	 */

	public void cols_insert(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		ColOperate colOperate = getJsonDataParameter(req, ColOperate.class);
		this.assembleData(req, resp,colOperate,OperatorConstant.colsinsert);
	}

	/**
	 * 删除列
	 * 
	 * @throws IOException
	 */

	public void cols_delete(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		ColOperate colOperate = getJsonDataParameter(req, ColOperate.class);
		this.assembleData(req, resp,colOperate,OperatorConstant.colsdelete);
	}

	/**
	 * 宽度调整
	 * 
	 * @throws IOException
	 */

	public void cols_width(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		ColWidth colWidth = getJsonDataParameter(req, ColWidth.class);
		this.assembleData(req, resp,colWidth,OperatorConstant.colswidth);
	}
	/**
	 * 列隐藏
	 * 
	 * @throws IOException
	 */

	public void cols_hide(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		ColOperate colOperate = getJsonDataParameter(req, ColOperate.class);
		this.assembleData(req, resp,colOperate,OperatorConstant.colshide);
	}
	/**
	 * 高度调整
	 * 
	 * @throws IOException
	 */

	public void rows_height(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		RowHeight rowHeight = getJsonDataParameter(req, RowHeight.class);
		this.assembleData(req, resp,rowHeight,OperatorConstant.rowsheight);
	}

	@Override
	public ModelAndView main(HttpServletRequest req, HttpServletResponse resp) {
		return null;
	}

}
