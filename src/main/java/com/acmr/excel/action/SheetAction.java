package com.acmr.excel.action;

import java.io.IOException;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import net.spy.memcached.MemcachedClient;

import org.springframework.web.servlet.ModelAndView;

import acmr.excel.pojo.ExcelBook;

import com.acmr.excel.action.excelbase.ExcelBaseAction;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.Frozen;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.RowLine;
import com.acmr.excel.model.complete.rows.ColOperate;
import com.acmr.excel.service.SheetService;
import com.acmr.excel.util.JsonReturn;
import com.acmr.excel.util.MemcacheUtil;
import com.acmr.helper.model.JSONReturnData;

/**
 * SHEET操作
 * 
 * @author jinhr
 *
 */

public class SheetAction extends ExcelBaseAction {


	/**
	 * 新建sheet
	 * 
	 * @throws IOException
	 */
	public void create(HttpServletRequest req, HttpServletResponse resp)
			throws IOException {
	}

	/**
	 * 修改sheet
	 * 
	 */
	public void update(HttpServletRequest req, HttpServletResponse resp)
			throws IOException {
	}

	/**
	 * 删除sheet
	 * 
	 * @throws IOException
	 */
	public void delete(HttpServletRequest req, HttpServletResponse resp)
			throws IOException {
	}

	/**
	 * 冻结
	 * 
	 * @throws Exception
	 */

	public void frozen(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		Frozen frozen = getJsonDataParameter(req, Frozen.class);
		this.assembleData(req, resp,frozen,OperatorConstant.frozen);
	}

	/**
	 * 取消冻结
	 * 
	 * @throws Exception
	 */

	public void unFrozen(HttpServletRequest req, HttpServletResponse resp) throws Exception {
//		String excelId = req.getParameter("excelId");
//		String sheetId = req.getParameter("sheetId");
//		ExcelBook excelBook = (ExcelBook) MemcacheUtil.get(excelId,memcachedClient);
//		JsonReturn data = new JsonReturn("");
//		if (excelBook != null) {
//			excelBook.getSheets().get(0).setFreeze(null);
//			MemcacheUtil.set(excelId, memcachedClient, excelBook);
//			data.setReturncode(200);
//		} else {
//			data.setReturncode(Constant.CACHE_INVALID_CODE);
//			data.setReturndata(Constant.CACHE_INVALID_MSG);
//		}
//		this.sendJson(resp, data);
		Frozen frozen = getJsonDataParameter(req, Frozen.class);
		this.assembleData(req, resp,frozen,OperatorConstant.unFrozen);
	}

	/**
	 * 增加线，用于初始化时向下滚动
	 */

	public void addrowline(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
//		String excelId = req.getParameter("excelId");
//		String sheetId = req.getParameter("sheetId");
//		String rowNum = req.getParameter("rowNum");
//		// ExcelBook excelBook = (ExcelBook)memcachedClient.get(excelId);
//		ExcelBook excelBook = (ExcelBook) MemcacheUtil.get(excelId,
//				memcachedClient);
//		JSONReturnData data = new JSONReturnData("");
//		if (excelBook != null) {
//			sheetService.addRowLine(excelBook.getSheets().get(0),
//					Integer.valueOf(rowNum));
//			MemcacheUtil.set(excelId, memcachedClient, excelBook);
//			data.setReturncode(200);
//		} else {
//			data.setReturncode(204);
//		}
//
//		this.sendJson(resp, data);
		RowLine rowLine = getJsonDataParameter(req, RowLine.class);
		this.assembleData(req, resp,rowLine,OperatorConstant.addRowLine);
	}
	/**
	 * 取消隐藏列
	 * 
	 * @throws Exception
	 */

	public void cols_cancelhide(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		ColOperate colOperate = getJsonDataParameter(req, ColOperate.class);
		this.assembleData(req, resp,colOperate,OperatorConstant.colhideCancel);
	}
	
	@Override
	public ModelAndView main(HttpServletRequest req, HttpServletResponse resp) {
		return null;
	}
}
