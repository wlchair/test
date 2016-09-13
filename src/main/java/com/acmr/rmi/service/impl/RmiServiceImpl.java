package com.acmr.rmi.service.impl;

import net.spy.memcached.MemcachedClient;
import acmr.excel.pojo.ExcelBook;

import com.acmr.excel.model.Constant;
import com.acmr.rmi.service.RmiService;

public class RmiServiceImpl implements RmiService {
	private MemcachedClient memcachedClient;

	public void setMemcachedClient(MemcachedClient memcachedClient) {
		this.memcachedClient = memcachedClient;
	}

	@Override
	public void saveExcelBook(String excelId, ExcelBook excelBook) {
		memcachedClient.set(excelId, Constant.MEMCACHED_EXP_TIME,excelBook);
		memcachedClient.set(excelId+"_ope", Constant.MEMCACHED_EXP_TIME, 0);
		System.out.println(memcachedClient.get(excelId+"_ope"));
	}

	@Override
	public ExcelBook getExcelBook(String excelId, int step) {
		ExcelBook excelBook = null;
		int curStep = (int) memcachedClient.get(excelId + "_ope");
		if (curStep == step) {
			excelBook = (ExcelBook) memcachedClient.get(excelId);
		} else {
			for (int i = 0; i < 100; i++) {
				int st = (int) memcachedClient.get(excelId + "_ope");
				if (step == st) {
					excelBook = (ExcelBook) memcachedClient.get(excelId);
				} else {
					try {
						Thread.sleep(100);
					} catch (InterruptedException e) {
						e.printStackTrace();
					}
				}
			}
		}
		return excelBook;
	}

	

}
