package com.acmr.excel.model;

import java.util.ArrayList;
import java.util.List;

public class Constant {
	public static final int SUCCESS_CODE = 200;
	public static final String SUCCESS_MSG = "操作成功";
	public static final int CACHE_INVALID_CODE = 300;
	public static final String CACHE_INVALID_MSG = "缓存失效";
	public static final int MEMCACHED_EXP_TIME = 60 * 60 * 24 * 25;
	public static final String queueName = "spreadsheet";

	/**
	 * 单元格高亮
	 */
	public static final String HIGHLIGHT = "highlight";

	public static List<String> accessControlAllowOriginList = new ArrayList<String>();
	
	static {
		accessControlAllowOriginList.add("http://localhost:4711");
		accessControlAllowOriginList.add("http://192.168.2.193:8080");
		accessControlAllowOriginList.add("http://192.168.2.241:8080");
		//accessControlAllowOriginList.add("http://localhost:8080");
		accessControlAllowOriginList.add("http://192.168.2.207:8080");
		accessControlAllowOriginList.add("http://192.168.2.16:8080");
		accessControlAllowOriginList.add("http://192.168.2.234:8080");
		accessControlAllowOriginList.add("http://192.168.1.241:8080");
		accessControlAllowOriginList.add("http://192.168.2.78:8080");
		accessControlAllowOriginList.add("http://192.168.2.73:8080");
		accessControlAllowOriginList.add("http://192.168.1.194:8080");
	}

}
