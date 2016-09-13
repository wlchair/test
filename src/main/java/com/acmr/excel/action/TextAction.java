package com.acmr.excel.action;

import java.io.IOException;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import net.spy.memcached.MemcachedClient;

import org.apache.poi.ss.format.CellFormat;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import acmr.excel.pojo.ExcelBook;

import com.acmr.core.util.string.StringUtil;
import com.acmr.excel.action.excelbase.ExcelBaseAction;
import com.acmr.excel.model.Cell;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.CellFormate.CellFormate;
import com.acmr.excel.model.comment.Comment;
import com.acmr.excel.service.HandleExcelService;
import com.acmr.excel.service.HandleExcelService.CellUpdateType;
import com.acmr.excel.util.MemcacheUtil;
import com.acmr.helper.model.JSONReturnData;
import com.acmr.mq.Model;
import com.acmr.mq.consumer.queue.QueueReceiver;
import com.acmr.mq.producer.queue.QueueSender;
import com.alibaba.fastjson.JSON;

@Controller
@RequestMapping("/text")
public class TextAction extends ExcelBaseAction {
	@Resource
	private QueueSender queueSender;
	@Resource
	private QueueReceiver queueReceiver;


	public ModelAndView main(HttpServletRequest req, HttpServletResponse resp) {
		return null;
	}

	/**
	 * 字号
	 * 
	 * @throws IOException
	 */

	public void font_size(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp, cell, OperatorConstant.fontsize);
	}

	/**
	 * 风格
	 * 
	 * @throws IOException
	 */

	public void font_family(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		// 接收参数，定义返回
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp, cell, OperatorConstant.fontfamily);
	}

	/**
	 * 粗细
	 * 
	 * @throws IOException
	 */

	public void font_weight(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		// 接收参数，定义返回
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp, cell, OperatorConstant.fontweight);
	}

	/**
	 * 斜体
	 * 
	 * @throws IOException
	 */

	public void font_italic(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp, cell, OperatorConstant.fontitalic);
	}

	/**
	 * 字体颜色
	 * 
	 * @throws IOException
	 */

	public void font_color(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		// 接收参数，定义返回
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp, cell, OperatorConstant.fontcolor);
	}
	/**
	 * 自动换行
	 * 
	 * @param req
	 * @param resp
	 * @throws Exception
	 */

	public void wordwrap(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp, cell, OperatorConstant.wordWrap);
	}
	/**
	 * 背景颜色
	 * 
	 * @throws IOException
	 */

	public void fill_bgcolor(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		// 接收参数，定义返回
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp, cell, OperatorConstant.fillbgcolor);
	}

	/**
	 * 编辑单元格中数据内容
	 * 
	 * @throws IOException
	 */

	public void data(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp, cell, OperatorConstant.textData);
	}
	/**
	 * 设置内容数据类型
	 * 
	 * @throws IOException
	 */

	public void data_format(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		CellFormate cellFormate = getJsonDataParameter(req, CellFormate.class);
		this.assembleData(req, resp, cellFormate, OperatorConstant.textDataformat);
	}

	/**
	 * 批量设置备注
	 * 
	 * @param req
	 * @param resp
	 * @throws Exception
	 */

	public void comment_set(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		Comment comment = getJsonDataParameter(req, Comment.class);
		this.assembleData(req, resp, comment, OperatorConstant.commentset);
	}
	
	
	/**
	 * 批量删除备注
	 * 
	 * @param req
	 * @param resp
	 * @throws Exception
	 */

	public void comment_del(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		comment_set(req, resp);
	}
	
	
	
	public void color_set(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp, cell, OperatorConstant.colorset);
	}
	
}
