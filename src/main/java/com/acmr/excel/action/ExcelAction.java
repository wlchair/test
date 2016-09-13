package com.acmr.excel.action;

import java.io.BufferedReader;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.List;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadPoolExecutor;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import net.spy.memcached.MemcachedClient;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;
import org.springframework.web.servlet.ModelAndView;

import acmr.excel.ExcelException;
import acmr.excel.pojo.Constants.XLSTYPE;
import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;

import com.acmr.core.util.string.StringUtil;
import com.acmr.excel.action.excelbase.ExcelBaseAction;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.OnlineExcel;
import com.acmr.excel.model.complete.CompleteExcel;
import com.acmr.excel.model.complete.ReturnParam;
import com.acmr.excel.model.complete.SpreadSheet;
import com.acmr.excel.model.position.OpenExcel;
import com.acmr.excel.model.position.Position;
import com.acmr.excel.service.DownLoadExcelService;
import com.acmr.excel.service.ExcelService;
import com.acmr.excel.service.HandleExcelService;
import com.acmr.excel.util.ExcelConst;
import com.acmr.excel.util.ExcelUtil;
import com.acmr.excel.util.FileUtil;
import com.acmr.excel.util.JsonReturn;
import com.acmr.excel.util.MemcacheUtil;
import com.acmr.excel.util.UUIDUtil;
import com.acmr.excel.util.UploadThread;
import com.acmr.helper.model.JSONReturnData;
import com.alibaba.fastjson.JSON;

/**
 * excel操作
 * 
 * @author jinhr
 */
// @Controller
// @RequestMapping("/excel")

public class ExcelAction extends ExcelBaseAction {
	private static Logger log = Logger.getLogger(ExcelAction.class); 
	// @Resource
	private HandleExcelService handleExcelService;
	// @Resource
	private DownLoadExcelService downLoadExcelService;
	// @Resource
	private ExcelService excelService;

	private MemcachedClient memcachedClient;
	

	public void setMemcachedClient(MemcachedClient memcachedClient) {
		this.memcachedClient = memcachedClient;
	}

	public void setHandleExcelService(HandleExcelService handleExcelService) {
		this.handleExcelService = handleExcelService;
	}

	public void setDownLoadExcelService(
			DownLoadExcelService downLoadExcelService) {
		this.downLoadExcelService = downLoadExcelService;
	}

	public void setExcelService(ExcelService excelService) {
		this.excelService = excelService;
	}
	
	
	// /**
	// * EXCEL新建
	// * @param name excel名称
	// * @return excel主页面
	// */
	// public ModelAndView create(HttpServletRequest req,HttpServletResponse
	// resp){
	// Excel excel = null;
	// try {
	// excel = getJsonDataParameter(req,Excel.class);
	// } catch (IOException e) {
	// // TODO Auto-generated catch block
	// e.printStackTrace();
	// }
	// String excelName = "";
	// if(excel != null){
	// excelName = excel.getName();
	// }
	// JSONReturnData data = new JSONReturnData("");
	// //生成excel唯一性id
	// String excelId = UUIDUtil.getUUID();
	// //创建基础excel,获取excel
	// CompleteExcel baseExcel = handleExcelService.createBasicExcel(excelId,
	// excelName);
	// String excelJson = JSON.toJSONString(baseExcel);
	// //存储excel到session
	// HttpSession session = req.getSession();
	// session.setAttribute(excelId, baseExcel);
	// //返回生成的excelId
	// data.setReturndata(excelId);
	// this.sendJson(resp, data);
	// return new ModelAndView("/index");
	// }
	// /**
	// * EXCEL名称修改
	// * @param excelId excelid
	// * @param name excel名称
	// * @throws IOException
	// */
	// public void update(HttpServletRequest req,HttpServletResponse resp)
	// throws IOException{
	// Excel excel = getJsonDataParameter(req,Excel.class);
	// //接收参数，定义返回
	// String excelName = excel.getName();
	// String excelId = excel.getExcelId();
	// JSONReturnData data = new JSONReturnData("");
	// //获取excel的完整对象
	// HttpSession session = req.getSession();
	// // String excelJson = "";
	// CompleteExcel cexcel = (CompleteExcel) session.getAttribute(excelId);
	// // if(!StringUtil.isEmpty(excelId)){
	// // excelJson = handleExcelService.getExcelJsonFromSessionById(excelId,
	// session);
	// // }
	// //
	// if(cexcel==null){
	// data.setReturncode(500);
	// data.setReturndata("没有查找到相应的excel信息");
	// }else{
	// // CompleteExcel cexcel = JSON.parseObject(excelJson,
	// CompleteExcel.class);
	// cexcel.setName(excelName);
	// // excelJson = JSON.toJSONString(cexcel);
	// //存储excel到session
	// session.setAttribute(excelId, cexcel);
	// }
	// //返回数据
	// this.sendJson(resp, data);
	// }
	// /**
	// * EXCEL 加载
	// * @param excelId id
	// * @param resp
	// * @throws IOException
	// */
	// public void reload(HttpServletRequest req,HttpServletResponse resp)
	// throws IOException{
	// Excel excel = getJsonDataParameter(req,Excel.class);
	// String excelId = req.getParameter("excelId");
	// HttpSession session = req.getSession();
	// String excelJson = "";
	// CompleteExcel cexcel = (CompleteExcel) session.getAttribute(excelId);
	// if(!StringUtil.isEmpty(excelId)){
	// excelJson = JSON.toJSONString(cexcel);
	// }
	// ////system.out.println(excelJson);
	// JSONReturnData data = new JSONReturnData("");
	// data.setReturndata(cexcel);
	// this.sendJson(resp, data);
	// }


	/**
	 * excel下载
	 */
	public void download(HttpServletRequest req, HttpServletResponse resp) {
		String excelId = req.getParameter("excelId");
		ExcelBook excelBook = (ExcelBook) memcachedClient.get(excelId);
		if (excelBook != null) {
			//excelService.changeHeightOrWidth(excelBook);
			try {
				OutputStream out = resp.getOutputStream();
				resp.setContentType("application/octet-stream");
				resp.setHeader("Content-disposition", "attachment;filename="+ URLEncoder.encode("模板" + ".xlsx", "utf-8"));
				excelBook.saveExcel(out, XLSTYPE.XLSX);
				out.flush();
				out.close();
			} catch (UnsupportedEncodingException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			} catch (ExcelException e) {
				e.printStackTrace();
			}
		}

	}

	/**
	 * 初始化excel页面
	 */
	@Override
	public ModelAndView main(HttpServletRequest req, HttpServletResponse resp) {
		String excelId = UUIDUtil.getUUID();
		ExcelBook excelBook = handleExcelService.createNewExcel(excelId);
		memcachedClient.set(excelId, Constant.MEMCACHED_EXP_TIME,excelBook);
		memcachedClient.set(excelId+"_ope", Constant.MEMCACHED_EXP_TIME, 0);
		log.info("初始化excel");
		// ExcelBook e = (ExcelBook)memcachedClient.get(excelId);
		// } <input type="hidden" id="excelId" value="(.*)"/>
		return new ModelAndView("/index").addObject("sheetId", "1").addObject("build", true).addObject("excelId", excelId);
	}

	/**
	 * 测试接口
	 * 
	 * @param req
	 * @param resp
	 * @throws InterruptedException 
	 */

	public void test(HttpServletRequest req, HttpServletResponse resp) throws InterruptedException {
		String flag = req.getParameter("flag");
		String excelId = UUIDUtil.getUUID();
		ExcelBook excelBook = handleExcelService.createNewExcel(excelId);
		ExcelSheet excelsheet = excelBook.getSheets().get(0);
		List<ExcelRow> rowList = excelsheet.getRows();
		for (int i = 0; i < 5; i++) {
			List<ExcelCell> excelList = rowList.get(i).getCells();
			for (int j = 0; j < 20; j++) {
				ExcelCell excelCell = excelList.get(j);
				if (excelCell == null) {
					excelCell = new ExcelCell();
				}
				excelCell.setText("回fff的痕迹卡的很金卡号地块和电视剧阿卡");
				excelCell.setValue("回到fffff的痕迹卡的很金卡号地块和电视剧阿卡");
				excelList.set(j, excelCell);
			}
		}
		excelsheet.MergedRegions(4, 4, 6, 6);
		memcachedClient.set(excelId, Constant.MEMCACHED_EXP_TIME,excelBook);
		memcachedClient.set(excelId+"_ope", Constant.MEMCACHED_EXP_TIME, 0);
		JsonReturn data = new JsonReturn("");
		data.setReturndata(excelId);
		this.sendJson(resp, data);
	}

	public void open(HttpServletRequest req, HttpServletResponse resp) throws InterruptedException {
		ExcelBook excelBook1 = (ExcelBook) MemcacheUtil.get("f3e2ebd8-f82e-4ba7-ae3c-70304e33232d", memcachedClient);
		System.out.println(JSON.toJSONString(excelBook1));
		//this.sendJson(resp, data);
	}
	
	/**
	 * 前台获得js文件
	 * 
	 * @param req
	 * @param resp
	 */

	public void getscript(HttpServletRequest req, HttpServletResponse resp) {
		String excelId = req.getParameter("excelId");
		String realPath = req.getSession().getServletContext().getRealPath("/");
		String jsString = readFile(realPath + "dist/my.js");
		String buildState = "window.SPREADSHEET_BUILD_STATE=";
		if (StringUtil.isEmpty(excelId)) {
			excelId = UUIDUtil.getUUID();
			ExcelBook excelBook = handleExcelService.createNewExcel(excelId);
			memcachedClient.set(excelId, Constant.MEMCACHED_EXP_TIME,excelBook);
			memcachedClient.set(excelId+"_ope", Constant.MEMCACHED_EXP_TIME, 0);
			buildState += "\"true\";";
			// ExcelBook e = (ExcelBook)memcachedClient.get(excelId);
		} else {
			buildState += "\"false\";";
		}
		String excelIdString = "window.SPREADSHEET_AUTHENTIC_KEY=\"" + excelId
				+ "\";";

		// } <input type="hidden" id="excelId" value="(.*)"/>
		resp.setContentType("application/javascript; charset=utf-8");
		try {
			resp.getWriter().print(
					excelIdString + "\r\n" + buildState + "\r\n" + jsString);
		} catch (IOException e1) {
			e1.printStackTrace();
		}

	}

	// private void test500LineData(CompleteExcel excel) {
	// SheetElement sheet = excel.getSpreadSheet().get(0).getSheet();
	// Map<String, Map<String, Integer>> alaisY =
	// sheet.getPosi().getStrandY().getAliasY();
	// for (int i = 0; i < 500; i++) {
	// for (int j = 0; j < 26; j++) {
	// Map<String, Integer> x = new HashMap<String, Integer>();
	// x.put(j + "", j);
	// alaisY.put(i + "", x);
	// }
	// }
	// for (int i = 0; i < 400; i++) {
	// Gly gly = new Gly();
	// gly.setAliasY("X");
	// gly.setHeight(600);
	// gly.setTop(300);
	// sheet.getGlY().add(gly);
	// }
	// for (int i = 1; i <= 13000; i++) {
	// sheet.getCells().add(new OneCell());
	// }
	// }

	/**
	 * 重新打开excel
	 * 
	 * @return
	 */
	public ModelAndView reOpen(HttpServletRequest req, HttpServletResponse resp) {
		String excelId = req.getParameter("excelId");
		//ExcelBook excelBook = (ExcelBook)ExcelBook.JSONParse(excel);
		ExcelBook excelBook = (ExcelBook)memcachedClient.get(excelId);
		memcachedClient.set(excelId, Constant.MEMCACHED_EXP_TIME, excelBook);
		return new ModelAndView("/index").addObject("excelId", excelId).addObject("sheetId", "1").addObject("build", false);
	}

	// public void getExcelParam(HttpServletRequest req, HttpServletResponse
	// resp) throws IOException {
	// String excelId = req.getParameter("excelId");
	// CompleteExcel excel = (CompleteExcel)
	// getSession(req).getAttribute(excelId);
	// if (excel != null) {
	// SheetElement sheet = excel.getSpreadSheet().get(0).getSheet();
	// int rowNum = sheet.getGlY().size();
	// int colNum = sheet.getGlX().size();
	// JsonReturn data = new JsonReturn("");
	// data.setRowNum(rowNum);
	// data.setColNum(colNum);
	// this.sendJson(resp, data);
	// }
	// }
	/**
	 * 通过像素动态加载excel
	 */

	public void openexcel(HttpServletRequest req, HttpServletResponse resp)
			throws IOException {
		// String excelId = req.getParameter("excelId");
		// //int sheetId = Integer.valueOf(req.getParameter("sheetId"));
		// int rowBegin = Integer.valueOf(req.getParameter("rowBegin"));
		// int rowEnd = Integer.valueOf(req.getParameter("rowEnd"));
		OpenExcel openExcel = getJsonDataParameter(req, OpenExcel.class);
		String excelId = openExcel.getExcelId();
		int memStep = (int)memcachedClient.get(excelId+"_ope");
		String curStep = req.getHeader("step");
		int cStep = 0;
		if(!StringUtil.isEmpty(curStep)){
			cStep = Integer.valueOf(curStep);
		}
		int rowBegin = openExcel.getRowBegin();
		int rowEnd = openExcel.getRowEnd();
		ExcelBook excelBook = (ExcelBook) memcachedClient.get(excelId);
		JsonReturn data = new JsonReturn("");
		if (cStep == memStep) {
			if (excelBook != null) {
				ExcelSheet excelSheet = excelBook.getSheets().get(0);
				ReturnParam returnParam = new ReturnParam();
				CompleteExcel excel = new CompleteExcel();
				SpreadSheet spreadSheet = new SpreadSheet();
				excel.getSpreadSheet().add(spreadSheet);
				spreadSheet = excelService.openExcel(spreadSheet, excelSheet,rowBegin, rowEnd, returnParam);
				data.setReturncode(Constant.SUCCESS_CODE);
				data.setReturndata(excel);
				data.setDataRowStartIndex(returnParam.getDataRowStartIndex());
			} else {
				data.setReturncode(Constant.CACHE_INVALID_CODE);
				data.setReturndata(Constant.CACHE_INVALID_MSG);
			}
		}else{
			for (int i = 0; i < 100; i++) {
				int mStep = (int)memcachedClient.get(excelId+"_ope");
				if(cStep == mStep){
					if (excelBook != null) {
						ExcelSheet excelSheet = excelBook.getSheets().get(0);
						ReturnParam returnParam = new ReturnParam();
						CompleteExcel excel = new CompleteExcel();
						SpreadSheet spreadSheet = new SpreadSheet();
						excel.getSpreadSheet().add(spreadSheet);
						spreadSheet = excelService.openExcel(spreadSheet, excelSheet,rowBegin, rowEnd, returnParam);
						data.setReturncode(Constant.SUCCESS_CODE);
						data.setReturndata(excel);
						data.setDataRowStartIndex(returnParam.getDataRowStartIndex());
					} else {
						data.setReturncode(Constant.CACHE_INVALID_CODE);
						data.setReturndata(Constant.CACHE_INVALID_MSG);
					}
				}else{
					try {
						Thread.sleep(100);
					} catch (InterruptedException e) {
						e.printStackTrace();
					}
				}
			}
		}
		
		
		// System.out.println("openexcel====================="+JSON.toJSONString(data));
		this.sendJson(resp, data);
	}

	/**
	 * 通过别名加载excel
	 */

	public void openExcelByAlais(HttpServletRequest req,HttpServletResponse resp) throws IOException {
		String excelId = req.getParameter("excelId");
		// int sheetId = Integer.valueOf(req.getParameter("sheetId"));
		String rowBegin = req.getParameter("rowBeginAlais");
		String rowEnd = req.getParameter("rowEndAlais");
		ExcelBook excelBook = (ExcelBook) memcachedClient.get(excelId);
		// Workbook workbook = this.mockWorkbook();
		// CompleteExcel excel = excelService.getExcel(workbook);
		JsonReturn data = new JsonReturn("");
		if (excelBook != null) {
			ExcelSheet excelSheet = excelBook.getSheets().get(0);
			ReturnParam returnParam = new ReturnParam();
			CompleteExcel excel = new CompleteExcel();
			SpreadSheet spreadSheet = new SpreadSheet();
			excel.getSpreadSheet().add(spreadSheet);
			spreadSheet = excelService.openExcelByAlais(spreadSheet,
					excelSheet, rowBegin, rowEnd, returnParam);
			data.setReturncode(Constant.SUCCESS_CODE);
			data.setReturndata(excel);
			data.setDataRowStartIndex(returnParam.getDataRowStartIndex());
		} else {
			data.setReturncode(Constant.CACHE_INVALID_CODE);
			data.setReturndata(Constant.CACHE_INVALID_MSG);
		}
		// System.out.println("openexcel====================="+JSON.toJSONString(data));
		this.sendJson(resp, data);
	}

	// public void openExcelByIndex(HttpServletRequest req,HttpServletResponse
	// resp) throws IOException {
	// String excelId = req.getParameter("excelId");
	// String rowBegin = req.getParameter("rowBegin");
	// String rowEnd = req.getParameter("rowEnd");
	// // String colEnd = req.getParameter("colEnd");
	// CompleteExcel excel =
	// (CompleteExcel)getSession(req).getAttribute(excelId);
	// Workbook workbook = this.mockWorkbook();
	// excel = excelService.getExcel(workbook);
	// if(excel != null){
	// excel = excelService.openExcelByIndex(excel,rowBegin,rowEnd);
	// }
	// JSONReturnData data = new JSONReturnData("");
	// String jsonData = JSON.toJSONString(excel);
	// ////system.out.println(jsonData);
	// data.setReturndata(excel);
	// this.sendJson(resp, data);
	// }
	// public void getExcelParam(HttpServletRequest req,HttpServletResponse
	// resp) throws IOException {
	// Workbook workbook = this.mockWorkbook();
	// String excelId = req.getParameter("excelId");
	// CompleteExcel excel = excelService.getExcel(workbook);
	// ////system.out.println(JSON.toJSONString(excel));
	// List<Gly> glys = excel.getSpreadSheet().get(0).getSheet().getGlY();
	// JsonReturn data = new JsonReturn("");
	// int rowNum = glys.size();
	// data.setRowNum(rowNum);
	// Gly gly = glys.get(rowNum-1);
	// int rowLength = gly.getTop() + gly.getHeight();
	// getSession(req).setAttribute(excelId, excel);
	// data.setRowLength(rowLength);
	// this.sendJson(resp, data);
	// }
	// public void getExcelParamByIndex(HttpServletRequest
	// req,HttpServletResponse resp) throws IOException {
	// Workbook workbook = this.mockWorkbook();
	// String excelId = req.getParameter("excelId");
	// CompleteExcel excel = excelService.getExcel(workbook);
	// ////system.out.println(JSON.toJSONString(excel));
	// List<Gly> glys = excel.getSpreadSheet().get(0).getSheet().getGlY();
	// getSession(req).setAttribute(excelId, excel);
	// JSONReturnData data = new JSONReturnData("");
	// data.setReturndata(glys.size());
	// this.sendJson(resp, data);
	// }
	// public void readExcel(HttpServletRequest req, HttpServletResponse resp)
	// throws IOException {
	// Workbook workbook = this.mockWorkbook();
	// CompleteExcel excel = excelService.getExcel(workbook);
	// ////system.out.println(JSON.toJSONString(excel));
	// }
	//

//	private Workbook mockWorkbook() {
//		File file = new File("d:/a2.xlsx");
//		InputStream is;
//		try {
//			is = new FileInputStream(file);
//			return new XSSFWorkbook(is);
//		} catch (FileNotFoundException e) {
//			e.printStackTrace();
//		} catch (IOException e) {
//			e.printStackTrace();
//		}
//		return null;
//	}

	private String readFile(String filepath) {
		String content = "";
		try {
			BufferedReader br = new BufferedReader(new FileReader(filepath));
			String str = null;
			StringBuffer buf = new StringBuffer();
			while ((str = br.readLine()) != null) {
				buf.append(str);
				buf.append("\r\n");
			}
			content = buf.toString();
			// System.out.println(content);
			br.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return content;
	}

	// public void upload(HttpServletRequest req, HttpServletResponse resp)
	// throws Exception{
	// //Workbook workbook = this.mockWorkbook();
	// List<MultipartFile> files = ((MultipartHttpServletRequest)
	// req).getFiles("file");
	// Workbook workbook =
	// ExcelUtil.createWorkbook(files.get(0).getInputStream());
	// CompleteExcel excel = excelService.getExcel(workbook);
	// String excelId = UUIDUtil.getUUID();
	// getSession(req).setAttribute(excelId, excel);
	// OnlineExcel onlineExcel = new OnlineExcel();
	// onlineExcel.setExcelId(excelId);
	// //System.out.println(JSONObject.toJSONString(excel));
	// onlineExcel.setJsonObject(JSONObject.toJSONString(excel));
	// onlineExcel.setName("测试上传"+System.currentTimeMillis());
	// JSONReturnData data = new JSONReturnData("");
	// try {
	// excelService.saveOrUpdateExcel(onlineExcel);
	// data.setReturncode(200);
	// data.setReturndata(excelId);
	// } catch (Exception e) {
	// data.setReturncode(202);
	// e.printStackTrace();
	// }
	// this.sendJson(resp, data);
	// }
	/**
	 * 上传excel
	 * 
	 * @param req
	 * @param resp
	 * @throws Exception
	 */

	public void upload(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		List<MultipartFile> files = ((MultipartHttpServletRequest) req).getFiles("file");
		ExcelBook excel = new ExcelBook();
		InputStream is = files.get(0).getInputStream();
		if (ExcelUtil.isExcel2003(is)) {
			excel.LoadExcel(files.get(0).getInputStream(), XLSTYPE.XLS);
		} else {
			excel.LoadExcel(files.get(0).getInputStream(), XLSTYPE.XLSX);
		}
		ExcelSheet excelSheet = excel.getSheets().get(0);
		List<ExcelRow> rowList = excelSheet.getRows();
		int rowSize = rowList.size();
		if (rowSize < 100) {
			for (int i = rowSize; i < 100; i++) {
				excelSheet.addRow();
			}
		}
		List<ExcelColumn> colList = excelSheet.getCols();
		int colSize = colList.size();
		if (colSize < 26) {
			for (int i = colSize; i < 26; i++) {
				excelSheet.addColumn();
			}
		}
		String excelId = UUIDUtil.getUUID();
		excelSheet.getExps().put("ifUpload", "true");
//		memcachedClient.set(excelId, 60 * 60 * 1, excel);
//		memcachedClient.set(excelId + "init", 60 * 60 * 1, excel);
		memcachedClient.set(excelId, Constant.MEMCACHED_EXP_TIME,excel);
		memcachedClient.set(excelId+"_ope", Constant.MEMCACHED_EXP_TIME, 0);
		JSONReturnData data = new JSONReturnData("");
		data.setReturncode(200);
		data.setReturndata(excelId);
		// ExcelBook excelBook = (ExcelBook)memcachedClient.get(excelId);
		// System.out.println("upload========================="+JSON.toJSONString(excelBook));
		this.sendJson(resp, data);
	}
//	public void save(HttpServletRequest req, HttpServletResponse resp)
//			throws Exception {
//		String excelId = req.getParameter("excelId");
//		CompleteExcel excel = (CompleteExcel) getSession(req).getAttribute(
//				excelId);
//		JSONReturnData data = new JSONReturnData("");
//		if (excel == null) {
//			data.setReturncode(ExcelConst.SESSION_INVALID);
//		} else {
//			Workbook workbook = downLoadExcelService.downloadExcel(excel);
//			data.setReturncode(ExcelConst.SUCCESS);
//			OutputStream out = resp.getOutputStream();
//			workbook.write(out);
//			data.setReturndata(out);
//		}
//		this.sendJson(resp, data);
//	}
	
	/**
	 * 上传完成之后的页面
	 * @param req
	 * @param resp
	 * @return
	 * @throws Exception
	 */
	
	public ModelAndView uploadComplete(HttpServletRequest req,
			HttpServletResponse resp) throws Exception {
		String excelId = req.getParameter("excelId");
		return new ModelAndView("/index").addObject("excelId", excelId);
	}

	/**
	 * 展示所有可以还原的excel
	 * 
	 * @return 列表页面
	 */
	public ModelAndView show(HttpServletRequest req, HttpServletResponse resp) {
		List<OnlineExcel> excels = excelService.getAllExcel();
		return new ModelAndView("/show").addObject("excels", excels);
	}

	/**
	 * 重新打开excel
	 * 
	 * @throws Exception
	 */
	// @RequestMapping(value="/position",method=RequestMethod.GET)

	public void position(HttpServletRequest req, HttpServletResponse resp) throws Exception {
//		String excelId = req.getParameter("excelId");
//		String height = req.getParameter("height");
		String excelId = req.getHeader("excelId");
		Position position = getJsonDataParameter(req, Position.class);
		int height = position.getContainerHeight();
//		if (StringUtil.isEmpty(height)) {
//			height = "800";
//		}
		ExcelBook excelBook = (ExcelBook) memcachedClient.get(excelId);
//		MemcacheUtil.set(excelId, memcachedClient, excelBook);
		//System.out.println(JSON.toJSONString(excelBook));
		ReturnParam returnParam = new ReturnParam();
		JsonReturn data = new JsonReturn("");
		CompleteExcel excel = new CompleteExcel();
		SpreadSheet spreadSheet = new SpreadSheet();
		if (excelBook != null) {
			ExcelSheet excelSheet = excelBook.getSheets().get(0);
			// if(excelSheet.getFreeze() != null){
			// SpreadSheet spreadSheet =
			// excelService.positionExcelWithFrozen(excel,
			// height,width,returnParam);
			// }else{
			spreadSheet = excelService.positionExcel(excelSheet, spreadSheet,height, returnParam);
			// }
		}
		excel.getSpreadSheet().add(spreadSheet);
		data.setReturncode(200);
		data.setMaxPixel(returnParam.getMaxPixel());
		data.setReturndata(excel);
		// data.setStartAlaisX(startAlais.getAlaisX());
		// data.setStartAlaisY(startAlais.getAlaisY());
		// data.setDataColStartIndex(returnParam.getDataColStartIndex());
		// data.setDataRowStartIndex(returnParam.getDataRowStartIndex());
		data.setDisplayColStartAlias(spreadSheet.getSheet().getGlX().get(0).getAliasX());
		data.setDisplayRowStartAlias(spreadSheet.getSheet().getGlY().get(0).getAliasY());
		memcachedClient.set(excelId+"_ope", Constant.MEMCACHED_EXP_TIME, 0);
		//System.out.println(JSON.toJSONString(data));
		this.sendJson(resp, data);
	}

	// /**
	// * 带冻结的定位还原
	// * @param req
	// * @param resp
	// * @throws Exception
	// */
	// public void positionWithFrozen(HttpServletRequest req,
	// HttpServletResponse resp) throws Exception{
	// String excelId = req.getParameter("excelId");
	// String height = req.getParameter("height");
	// String width = req.getParameter("width");
	// height = "800";
	// width = "300";
	// CompleteExcel excel =
	// (CompleteExcel)getSession(req).getAttribute(excelId);
	// Workbook workbook = this.mockWorkbook();
	// excel = excelService.getExcel(workbook);
	// SheetElement sheet = excel.getSpreadSheet().get(0).getSheet();
	// StartAlais startAlaisTest = sheet.getStartAlais();
	// sheet.getFrozen().setColIndex("3");
	// sheet.getFrozen().setRowIndex("22");
	// sheet.getFrozen().setDisplayAreaStartAlaisX("3");
	// sheet.getFrozen().setDisplayAreaStartAlaisY("22");
	// sheet.getFrozen().setState("1");
	// int glyLength = sheet.getGlY().size();
	// Gly gly = sheet.getGlY().get(glyLength-1);
	// int maxPixel = gly.getHeight() + gly.getTop();
	// startAlaisTest.setAlaisX("1");
	// startAlaisTest.setAlaisY("19");
	// ReturnParam returnParam = new ReturnParam();
	// if(excel != null){
	// excel = excelService.positionExcelWithFrozen(excel,
	// height,width,returnParam);
	// }
	// JsonReturn data = new JsonReturn("");
	// data.setReturncode(200);
	// data.setReturndata(excel);
	// // data.setxStartAlaisIndex(returnParam.getxStartAlaisIndex());
	// // data.setxEndAlaisIndex(returnParam.getxEndAlaisIndex());
	// // data.setyStartAlaisIndex(returnParam.getyStartAlaisIndex());
	// // data.setyEndAlaisIndex(returnParam.getyEndAlaisIndex());
	// data.setDataColStartIndex(returnParam.getDataColStartIndex());
	// data.setDataRowStartIndex(returnParam.getDataRowStartIndex());
	// data.setDisplayColStartAlias(returnParam.getDisplayColStartAlias());
	// data.setDisplayRowStartAlias(returnParam.getDisplayRowStartAlias());
	// data.setMaxPixel(maxPixel);
	// String jsonData = JSON.toJSONString(data);
	// ////system.out.println(jsonData);
	// this.sendJson(resp, data);
	// }
	/**
	 * 保存excel(关闭浏览器时的操作)
	 * 
	 * @throws Exception
	 */

	public void save(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		String excelId = req.getParameter("excelId");
		// String startX = req.getParameter("startX");
		// String startY = req.getParameter("startY");
		Thread.sleep(3000);
		ExcelBook excelBook = (ExcelBook) memcachedClient.get(excelId);
		JsonReturn data = new JsonReturn("");
		if (excelBook != null) {
			OnlineExcel olExcel = new OnlineExcel();
			olExcel.setExcelId(excelId);
			//olExcel.setExcelObject(ExcelBook.JSONString());
			excelService.saveOrUpdateExcel(olExcel);
			memcachedClient.set(excelId, Constant.MEMCACHED_EXP_TIME,excelBook);
			data.setReturncode(200);
		}
		this.sendJson(resp, data);
	}
	public void close(HttpServletRequest req, HttpServletResponse resp){
	}
	// public void test(HttpServletRequest request,HttpServletResponse
	// response){
	// this.memcachedClient.add("a", 7200, "aaaaaaaaaaaa");
	// String result = this.memcachedClient.get("a").toString();
	// System.out.println(result);
	// }

	// public void uploadBigFile(HttpServletRequest req, HttpServletResponse
	// resp){
	// List<MultipartFile> files = ((MultipartHttpServletRequest)
	// req).getFiles("file");
	// if (!files.isEmpty() && files.size() > 0) {
	// ThreadPoolExecutor threadPool = (ThreadPoolExecutor)
	// Executors.newCachedThreadPool();
	// for (int i = 0; i < files.size(); i++) {
	// MultipartFile file = files.get(i);
	// String partFileName = file.getName() + "." + (i+1) + ".part";
	// try {
	// threadPool.execute(new UploadThread(partFileName, file.getBytes()));
	// } catch (IOException e) {
	// e.printStackTrace();
	// }
	// }
	// }
	// }

	/**
	 * 大文件上传
	 * 
	 * @param req
	 * @param resp
	 */

	public void uploadBigFile(HttpServletRequest req, HttpServletResponse resp) {
		InputStream is;
		try {
			is = req.getInputStream();
			byte[] bytes = FileUtil.toByteArray(is);
			String name = req.getParameter("fname");
			ThreadPoolExecutor threadPool = (ThreadPoolExecutor) Executors
					.newCachedThreadPool();
			String partFileName = name + "." + System.currentTimeMillis()
					+ ".part";
			threadPool.execute(new UploadThread(partFileName, bytes));
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * 合并文件
	 * 
	 * @param req
	 * @param resp
	 */

	public void mergeFile(HttpServletRequest req, HttpServletResponse resp) {
		FileUtil fileUtil = new FileUtil();
		int blockFileSize = 1024 * 1024 * 10;
		String name = req.getParameter("fname");
		try {
			fileUtil.mergePartFiles(ExcelUtil.currentWorkDir, ".part",
					blockFileSize, ExcelUtil.currentWorkDir + name);
		} catch (IOException e) {
			e.printStackTrace();
		}
		JSONReturnData data = new JSONReturnData("");
		data.setReturncode(200);
		String address = "d:\\temp\\" + name;
		data.setReturndata(address);
		this.sendJson(resp, data);
	}

	/**
	 * 大文件下载
	 * 
	 * @param req
	 * @param resp
	 * @throws IOException
	 */

	public void downloadBigFile(HttpServletRequest req, HttpServletResponse resp)
			throws IOException {
		String fileName = req.getParameter("fname");
		InputStream is = new FileInputStream("d:\\temp\\" + fileName);
		resp.reset();
		resp.setContentType("application/pdf");
		resp.setHeader("Pragma", "public");
		resp.setHeader("Cache-Control", "max-age=30");
		resp.setHeader("Content-disposition", "inline;filename=" + fileName);
		ServletOutputStream out = resp.getOutputStream();
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		byte[] bytes = FileUtil.toByteArray(is);
		try {
			if (null != bytes) {
				bos.write(bytes);
				resp.setContentLength(bos.size());
				bos.writeTo(out);
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			out.close();
			out.flush();
			bos.close();
			bos.flush();
		}
	}

}
