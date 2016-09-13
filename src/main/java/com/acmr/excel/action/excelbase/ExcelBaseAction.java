package com.acmr.excel.action.excelbase;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import net.spy.memcached.MemcachedClient;

import org.apache.log4j.Logger;
import org.springframework.web.servlet.ModelAndView;

import com.acmr.core.action.BaseAction;
import com.acmr.core.util.string.StringUtil;
import com.acmr.excel.action.ExcelAction;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.complete.CompleteExcel;
import com.acmr.helper.model.JSONReturnData;
import com.acmr.mq.Model;
import com.acmr.mq.producer.queue.QueueSender;
import com.alibaba.fastjson.JSON;


/**
 * excel base action
 * 
 * @author caosl
 */

public class ExcelBaseAction extends BaseAction {
	private static Logger logger = Logger.getLogger(ExcelAction.class);
	@Resource
	private QueueSender queueSender;
	@Override
	public ModelAndView main(HttpServletRequest req, HttpServletResponse resp) {
		return null;
	}

	// /**
	// * 获取decode解码后的参数对象
	// * @param <T>
	// * @param req
	// * @param key
	// * @return <T> t
	// * @throws IllegalAccessException
	// * @throws InstantiationException
	// */
	// public <T> T getJsonDataParameter(HttpServletRequest req, Class<T> t){
	// String value = req.getParameter(ExcelConst.PARAMETER_HEAD_STRING);
	// if (StringUtil.isEmpty(value)) {
	// try {
	// return t.newInstance();
	// } catch (InstantiationException | IllegalAccessException e) {
	// e.printStackTrace();
	// }
	// }
	// if (req.getMethod().toLowerCase().equals("get")) {
	// try {
	// value = new String(value.replaceAll("%", "###").getBytes(
	// "iso-8859-1"), "utf-8");
	// value = java.net.URLDecoder.decode(value, "utf-8");
	// } catch (UnsupportedEncodingException e1) {
	// }
	// }
	// return JSON.parseObject(value, t);
	// }
	/**
	 * session缓存
	 * 
	 * @param req
	 *            request对象
	 * @return session对象
	 */

	protected HttpSession getSession(HttpServletRequest req) {
		return req.getSession();
	}

	/**
	 * 把对象放入缓存
	 * 
	 * @param req
	 *            request对象
	 * @param memcachedClient
	 *            memcached对象
	 * @param key
	 *            键
	 * @param excel
	 *            值
	 */

	protected void setCache(HttpServletRequest req,
			MemcachedClient memcachedClient, String key, CompleteExcel excel) {
		memcachedClient.set(key, 60 * 60 * 24, excel);
		req.getSession().setAttribute(key, excel);
	}

	/**
	 * 把对象从缓存中取出
	 * 
	 * @param req
	 *            request对象
	 * @param memcachedClient
	 *            memcached对象
	 * @param key
	 *            键
	 * @return excel对象
	 */

	protected CompleteExcel getCache(HttpServletRequest req,
			MemcachedClient memcachedClient, String key) {
		return (CompleteExcel) memcachedClient.get(key);
		// return (CompleteExcel)req.getSession().setAttribute(key);
	}

	/**
	 * 把json串转化为实际对象
	 * 
	 * @param request
	 *            request对象
	 * @param t
	 *            泛型对象
	 * @return 实际对象
	 * @throws IOException
	 */

	protected <T> T getJsonDataParameter(HttpServletRequest request,
			Class<T> clazz) throws IOException {
		String body = null;
		StringBuilder stringBuilder = new StringBuilder();
		BufferedReader bufferedReader = null;
		try {
			InputStream inputStream = request.getInputStream();
			if (inputStream != null) {
				bufferedReader = new BufferedReader(new InputStreamReader(
						inputStream));
				char[] charBuffer = new char[128];
				int bytesRead = -1;
				while ((bytesRead = bufferedReader.read(charBuffer)) > 0) {
					stringBuilder.append(charBuffer, 0, bytesRead);
				}
			} else {
				stringBuilder.append("");
			}
		} catch (IOException ex) {
			throw ex;
		} finally {
			if (bufferedReader != null) {
				try {
					bufferedReader.close();
				} catch (IOException ex) {
					throw ex;
				}
			}
		}
		body = stringBuilder.toString();
		if (StringUtil.isEmpty(body)) {
			try {
				return clazz.newInstance();
			} catch (InstantiationException | IllegalAccessException e) {
				e.printStackTrace();
			}
		}
		return JSON.parseObject(body, clazz);
	}

	protected String getBody(HttpServletRequest request) throws IOException {
		String body = null;
		StringBuilder stringBuilder = new StringBuilder();
		BufferedReader bufferedReader = null;
		try {
			InputStream inputStream = request.getInputStream();
			if (inputStream != null) {
				bufferedReader = new BufferedReader(new InputStreamReader(
						inputStream));
				char[] charBuffer = new char[128];
				int bytesRead = -1;
				while ((bytesRead = bufferedReader.read(charBuffer)) > 0) {
					stringBuilder.append(charBuffer, 0, bytesRead);
				}
			} else {
				stringBuilder.append("");
			}
		} catch (IOException ex) {
			throw ex;
		} finally {
			if (bufferedReader != null) {
				try {
					bufferedReader.close();
				} catch (IOException ex) {
					throw ex;
				}
			}
		}

		body = stringBuilder.toString();
		return body;
	}

	@Override
	protected void sendJson(HttpServletResponse resp, Object data) {
		// resp.setHeader("Access-Control-Allow-Origin", "*");
		super.sendJson(resp, data);
	}
	protected void assembleData(HttpServletRequest req,HttpServletResponse resp,Object object,int reqPath){
		JSONReturnData data = new JSONReturnData("");
		String step = req.getHeader("step");
		String excelId = req.getHeader("excelId");
		if(StringUtil.isEmpty(step)){
			step = "1";
		}
		Model model = new Model();
		int index = Integer.valueOf(step);
		model.setStep(index);
		model.setReqPath(reqPath);
		model.setObject(object);
		model.setExcelId(excelId);
		logger.info("**********发送excelId:"+excelId+"====step:"+step+"===reqPath:"+reqPath);	
		queueSender.send(Constant.queueName, model);
		data.setReturncode(Constant.SUCCESS_CODE);
		data.setReturndata(true);
		this.sendJson(resp, data);
	}
	protected void assemblePasteData(HttpServletRequest req,HttpServletResponse resp){
		JSONReturnData data = new JSONReturnData("");
		String step = req.getHeader("step");
		String excelId = req.getHeader("excelId");
		if(StringUtil.isEmpty(step)){
			step = "1";
		}
		Model model = new Model();
		int index = Integer.valueOf(step);
		model.setStep(index);
		model.setReqPath(29);
		model.setExcelId(excelId);
			//Thread.sleep(index*100);
		queueSender.send(Constant.queueName, model);
		data.setReturncode(Constant.SUCCESS_CODE);
		data.setReturndata(false);
		this.sendJson(resp, data);
	}
}
