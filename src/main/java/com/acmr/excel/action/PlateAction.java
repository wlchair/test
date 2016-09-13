
package com.acmr.excel.action;

import java.io.IOException;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import net.spy.memcached.MemcachedClient;

import org.springframework.web.servlet.ModelAndView;

import acmr.excel.pojo.ExcelBook;

import com.acmr.excel.action.excelbase.ExcelBaseAction;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.Paste;
import com.acmr.excel.model.copy.Copy;
import com.acmr.excel.service.PasteService;
import com.acmr.excel.util.MemcacheUtil;
import com.acmr.helper.model.JSONReturnData;

/**
 * 剪切板功能
 * @author jinhr
 *
 */
public class PlateAction extends ExcelBaseAction{
	private MemcachedClient memcachedClient;
	private PasteService pasteService; 
	
	public void setPasteService(PasteService pasteService) {
		this.pasteService = pasteService;
	}
	
	public void setMemcachedClient(MemcachedClient memcachedClient) {
		this.memcachedClient = memcachedClient;
	}
	
	@Override
	public ModelAndView main(HttpServletRequest req, HttpServletResponse resp) {
		return null;
	}
	/**
	 * 外部粘贴
	 * @throws IOException
	 */
	
	public void paste(HttpServletRequest req,HttpServletResponse resp) throws Exception{
		Paste paste = getJsonDataParameter(req, Paste.class);
		ExcelBook excelBook = (ExcelBook)memcachedClient.get(paste.getExcelId());
		boolean isAblePasteResult = pasteService.isAblePaste(paste, excelBook);
		if(isAblePasteResult){
			this.assembleData(req, resp, paste, OperatorConstant.paste);
		}else{
			this.assemblePasteData(req, resp);
		}
		
	}
	/**
	 * 是否可以粘贴
	 * @throws IOException
	 */
	
	public void isAblePaste(HttpServletRequest req,HttpServletResponse resp) throws Exception{
//		IsAblePaste isAblePaste = getJsonDataParameter(req, IsAblePaste.class);
//		String excelId = isAblePaste.getExcelId();
//		ExcelBook excelBook = (ExcelBook) MemcacheUtil.get(excelId, memcachedClient);
//		JSONReturnData data = new JSONReturnData("");
//		boolean isAblePasteResult = pasteService.isAblePaste(isAblePaste, excelBook);
//		MemcacheUtil.set(excelId, memcachedClient, excelBook);
//		data.setReturncode(200);
//		data.setReturndata(isAblePasteResult);
//		this.sendJson(resp, data);
		
	}
	/**
	 * 内部复制粘贴
	 * @throws IOException
	 */
	
	public void copy(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		Copy copy = getJsonDataParameter(req, Copy.class);
		ExcelBook excelBook = (ExcelBook)memcachedClient.get(copy.getExcelId());
		boolean isAblePasteResult = pasteService.isCopyPaste(copy, excelBook);
		if(isAblePasteResult){
			this.assembleData(req, resp, copy, OperatorConstant.copy);
		}else{
			this.assemblePasteData(req, resp);
		}
	}
	/**
	 * 剪切粘贴
	 * @throws IOException
	 */
	
	public void cut(HttpServletRequest req,HttpServletResponse resp) throws Exception{
		Copy copy = getJsonDataParameter(req, Copy.class);
		ExcelBook excelBook = (ExcelBook)memcachedClient.get(copy.getExcelId());
		boolean isAblePasteResult = pasteService.isCopyPaste(copy, excelBook);
		if(isAblePasteResult){
			this.assembleData(req, resp, copy, OperatorConstant.cut);
		}else{
			this.assemblePasteData(req, resp);
		}
	}
}
