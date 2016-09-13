package com.acmr.excel.util;

import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import acmr.excel.pojo.Constants.CELLTYPE;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelColor;
import acmr.excel.pojo.ExcelFormat;

import com.acmr.excel.model.complete.Content;
import com.acmr.excel.model.complete.CustomProp;

public class CellFormateUtil {
	/**
	 * 文本类型
	 * @param excelCell
	 */
	public static void setText(ExcelCell excelCell) {
		//excelCell.setShowText(excelCell.getText());
		excelCell.setType(CELLTYPE.STRING);
		excelCell.getCellstyle().setDataformat("@");
		excelCell.setValue(excelCell.getText());
	}
	/**
	 * 常规类型
	 * @param excelCell
	 */
	public static void setGeneral(ExcelCell excelCell) {
		String text = excelCell.getText();
		//excelCell.setShowText(text);
		if(isNumeric(text)){
			excelCell.setType(CELLTYPE.NUMERIC);
		}else{
			excelCell.setType(CELLTYPE.STRING);
		}
		excelCell.getCellstyle().setDataformat("General");
	}
	/**
	 * 数字类型
	 * @param excelCell
	 * @param decimalPoint
	 * @param thousandPoint
	 */
	public static void setNumber(ExcelCell excelCell,int decimalPoint,boolean thousandPoint){
		String text = excelCell.getText();
		if (isNumeric(text)) {
			excelCell.setValue(Double.valueOf(text));
			//小数位
			text = setDecimalPoint(decimalPoint, text);
			//千分位
			text = setThousandPoint(thousandPoint, text);
			//excelCell.setShowText(text);
			excelCell.setType(CELLTYPE.NUMERIC);
			excelCell.getCellstyle().setDataformat(getNumDataFormate(decimalPoint, thousandPoint));
			excelCell.getExps().put("decimalPoint", decimalPoint+"");
			excelCell.getExps().put("thousandPoint", thousandPoint+"");
		}
	}
	/**
	 * 日期类型
	 * @param excelCell
	 */
	public static void setTime(ExcelCell excelCell,String dataFormate) {
		String text = excelCell.getText();
		String textFormate = "";
		if(text.contains("年") && text.contains("月") && text.contains("日")){
			textFormate = "yyyy年MM月dd日";
		}else if(text.contains("年") && text.contains("月") && !text.contains("日")){
			textFormate = "yyyy年MM月";
		}else if(text.contains("/")){
			textFormate = "yyyy/MM/dd";
		}
		String outputFormate = null;
		switch (dataFormate) {
		case "yyyy年MM月dd日":
			outputFormate = "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy";
			break;
		case "yyyy年MM月":
			outputFormate = "yyyy\"年\"m\"月\";@";
			break;
		case "yyyy/MM/dd":
			outputFormate = "m/d/yy";
			break;

		default:
			break;
		}
		SimpleDateFormat format = null;
		switch (textFormate) {
		case "yyyy年MM月dd日":
		    format = new SimpleDateFormat("yyyy年MM月dd日");
		    try {
		    	format.setLenient(false);
				Date date = format.parse(text);
				format = new SimpleDateFormat(dataFormate);
				String newDate = format.format(date);
				excelCell.setValue(date);
				//excelCell.setShowText(newDate);
				excelCell.setType(CELLTYPE.DATE);
				if(outputFormate != null){ 
					excelCell.getCellstyle().setDataformat(outputFormate);
				}
				excelCell.getExps().put("dataFormate", dataFormate);
			} catch (ParseException e) {
				//e.printStackTrace();
			}
			break;
		case "yyyy年MM月":
		    format = new SimpleDateFormat("yyyy年MM月");
		    try {
		    	format.setLenient(false);
				Date date = format.parse(text);
				format = new SimpleDateFormat(dataFormate);
				String newDate = format.format(date);
				excelCell.setValue(date);
				//excelCell.setShowText(newDate);
				excelCell.setType(CELLTYPE.DATE);
				if(outputFormate != null){
					excelCell.getCellstyle().setDataformat(outputFormate);
				}
				excelCell.getExps().put("dataFormate", dataFormate);
			} catch (ParseException e) {
				//e.printStackTrace();
			}
			break;
		case "yyyy/MM/dd":
		    format = new SimpleDateFormat("yyyy/MM/dd");
		    try {
		    	format.setLenient(false);
				Date date = format.parse(text);
				format = new SimpleDateFormat(dataFormate);
				String newDate = format.format(date);
				excelCell.setValue(date);
				//excelCell.setShowText(newDate);
				excelCell.setType(CELLTYPE.DATE);
				if(outputFormate != null){
					excelCell.getCellstyle().setDataformat(outputFormate);
				}
				excelCell.getExps().put("dataFormate", dataFormate);
			} catch (ParseException e) {
				//e.printStackTrace();
			}
			break;

		default:
			break;
		}
	}
	/**
	 * 货币类型
	 * @param excelCell
	 * @param decimalPoint
	 * @param currencySymbol
	 */
	public static void setCurrency(ExcelCell excelCell, int decimalPoint, String currencySymbol) {
		String text = excelCell.getText();
		if (isNumeric(text)) {
			excelCell.setValue(Double.valueOf(text));
			// 小数位
			text = setDecimalPoint(decimalPoint, text);
			// 千分位
			text = setThousandPoint(true, text);
			//excelCell.setShowText(currencySymbol + text);
			excelCell.setType(CELLTYPE.NUMERIC);
			excelCell.getCellstyle().setDataformat(getCurrencyDataFormate(decimalPoint, currencySymbol));
			excelCell.getExps().put("decimalPoint", decimalPoint+"");
			excelCell.getExps().put("currencySymbol", currencySymbol);
		}
	}
	/**
	 * 百分比
	 * @param excelCell
	 * @param decimalPoint
	 */
	public static void setPercent(ExcelCell excelCell, int decimalPoint) {
		String text = excelCell.getText();
		if (isNumeric(text)) {
			double num = Double.valueOf(text);
			excelCell.setValue(num);
			num *= 100;
			text = setDecimalPoint(decimalPoint, num + "");
			//excelCell.setShowText(text + "%");
			excelCell.setType(CELLTYPE.NUMERIC);
			excelCell.getCellstyle().setDataformat(getPercentDataFormate(decimalPoint));
		}
	}
	
	
	/**
	 * 日期校验
	 * @param date
	 * @param pattern
	 * @return
	 */
	public static Date getDate(String date,String pattern) {
		//boolean convertSuccess = true;
		Date retDate = null;
		SimpleDateFormat format = new SimpleDateFormat(pattern);
		try {
			format.setLenient(false);
			retDate = format.parse(date);
		} catch (ParseException e) {
			//convertSuccess = false;
		}
		return retDate;
	}
	
	/**
	 * 设置数字格式
	 * @param decimalPoint
	 * @param thousandPoint
	 * @return
	 */
	public static String getNumDataFormate(int decimalPoint,boolean thousandPoint) {
		String dataFormate = "";
		if(thousandPoint){
			dataFormate += "#,##";
		}
		dataFormate += "0";
		for (int i = 0; i < decimalPoint; i++) {
			if (i == 0) {
				dataFormate += ".";
			}
			dataFormate += "0";
		}
		dataFormate += "_);\\(" + dataFormate + "\\)";
		return dataFormate;
	}
	/**
	 * 设置货币格式
	 * @param decimalPoint
	 * @return
	 */
	public static String getCurrencyDataFormate(int decimalPoint,String currencySymbol) {
		String dataFormate = "\""+currencySymbol + "\"#,##0";
		for (int i = 0; i < decimalPoint; i++) {
			if (i == 0) {
				dataFormate += ".";
			}
			dataFormate += "0";
		}
		dataFormate += "_);\\(" + dataFormate + "\\)";
		return dataFormate;
	}

	/**
	 * 设置百分比格式
	 * 
	 * @param decimalPoint
	 * @return
	 */
	public static String getPercentDataFormate(int decimalPoint) {
		String dataFormate = "0";
		for (int i = 0; i < decimalPoint; i++) {
			if (i == 0) {
				dataFormate += ".";
			}
			dataFormate += "0";
		}
		dataFormate += "%";
		return dataFormate;
	}
	/**
	 * 设置小数位
	 * @param decimalPoint
	 * @param text
	 * @return
	 */
	public static String setDecimalPoint(int decimalPoint, String text) {
		if (decimalPoint < 0) {
			return null;
		}
		String temp = "#";
		for (int i = 0; i < decimalPoint; i++) {
			if(i == 0){
				temp += ".";
			}
			temp += "0";
		}
		DecimalFormat df = new DecimalFormat(temp);
		return df.format(Double.valueOf(text));
	}
	
	/**
	 * 设置千分位
	 * @param thousandPoint
	 * @param text
	 * @return
	 */
	public static String setThousandPoint(boolean thousandPoint,String text){
		String retVal = null;
		NumberFormat numberFormat = NumberFormat.getNumberInstance();
		if(thousandPoint){
			//千分位
			numberFormat.setGroupingUsed(true);
		}else{
			numberFormat.setGroupingUsed(false);
		}
		if(text.indexOf(".") < 0){
			//整数
			retVal = numberFormat.format(Long.valueOf(text));
		}else{
			int textLength = text.split("\\.")[1].length();
			//小数
			retVal = numberFormat.format(Double.valueOf(text));
			String[] news = retVal.split("\\.");
			int retLength = 0;
			if(news != null && news.length > 1){
				retLength = news[1].length();
			}else{
				retVal += ".";
			}
			if(textLength > retLength){
				for(int i = retLength;i< textLength;i++){
					retVal += 0;
				}
			}
		}
		return retVal;
	}
	
	/**
	 * 判断是否为数字
	 * @param str
	 * @return
	 */
	public static boolean isNumeric(String str) {
		//[-+]?\\d*\\.?\\d+, ^[0-9]+(\\.[0-9]+){0,1}$
		//Pattern pattern = Pattern.compile("-?[0-9]+[.]?[0-9]+");
		Pattern pattern = Pattern.compile("[-+]?\\d*\\.?\\d+");
		Matcher isNum = pattern.matcher(str);
		if (!isNum.matches()) {
			return false;
		}
		return true;
		//return NumberUtils.isNumber(str);
	}
	/**
	 * 自动识别
	 * @param content
	 */
	
	public static void autoRecognise(String content,ExcelCell excelCell){
		String pattern1 = "yyyy年MM月dd日"; 
		String pattern2 = "yyyy年MM月";
		String pattern3 = "yyyy/MM/dd";
		Date d1 = getDate(content, pattern1);
		Date d2 = getDate(content, pattern2);
		Date d3 = getDate(content, pattern3);
		if(isNumeric(content)){
			content = getPrettyNumber(content);
			excelCell.setType(CELLTYPE.NUMERIC);
			excelCell.getCellstyle().setDataformat("General");
			excelCell.setValue(Double.valueOf(content));
			//excelCell.setShowText(content);
		}else if(d1 != null){
			excelCell.setValue(d1);
			excelCell.setType(CELLTYPE.DATE);
			excelCell.getCellstyle().setDataformat("[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy");
			//excelCell.setShowText(content);
		}else if(d2 != null){
			excelCell.setValue(d2);
			excelCell.setType(CELLTYPE.DATE);
			excelCell.getCellstyle().setDataformat("yyyy\"年\"m\"月\";@");
			//excelCell.setShowText(content);
		}else if(d3 != null){
			excelCell.setValue(d3);
			excelCell.setType(CELLTYPE.DATE);
			excelCell.getCellstyle().setDataformat("m/d/yy");
			//excelCell.setShowText(content);
		}
	}
	
	public static void setShowText(ExcelCell excelCell,Content content,CustomProp customProp){
		if(CELLTYPE.BLANK == excelCell.getType()){
			content.setDisplayTexts("");
			return;
		}
		String dataFormate = excelCell.getCellstyle().getDataformat();
		String text = excelCell.getText();
		SimpleDateFormat sdf = null;
		switch (dataFormate) {
		case "General":
			customProp.setFormat("normal");
			content.setDisplayTexts(text);
			break;
		case "@":
			customProp.setFormat("text");
			content.setDisplayTexts(text);
			break;
		case "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy":
			sdf = new SimpleDateFormat("yyyy年MM月dd日");
			content.setDisplayTexts(sdf.format(excelCell.getValue()));
			content.setTexts(content.getDisplayTexts());
			customProp.setFormat("date");
			customProp.setDateFormat("yyyy年MM月dd日");
			if(CELLTYPE.STRING == excelCell.getType()){
				customProp.setIsValid(false);
			}
			break;
		case "yyyy\"年\"m\"月\";@":
			sdf = new SimpleDateFormat("yyyy年MM月");
			content.setDisplayTexts(sdf.format(excelCell.getValue()));
			customProp.setFormat("date");
			customProp.setDateFormat("yyyy年MM月");
			content.setTexts(content.getDisplayTexts());
			if(CELLTYPE.STRING == excelCell.getType()){
				customProp.setIsValid(false);
			}
			break;
		case "m/d/yy":
			sdf = new SimpleDateFormat("yyyy/MM/dd");
			content.setDisplayTexts(sdf.format(excelCell.getValue()));
			customProp.setFormat("date");
			customProp.setDateFormat("yyyy/MM/dd");
			content.setTexts(content.getDisplayTexts());
			if(CELLTYPE.STRING == excelCell.getType()){
				customProp.setIsValid(false);
			}
			break;
		case "0_);\\(0\\)":
			//整数
			customProp.setFormat("number");
			content.setDisplayTexts(text);
			customProp.setDecimal(0);
			customProp.setThousands(false);
			if(CELLTYPE.STRING == excelCell.getType()){
				customProp.setIsValid(false);
			}
			break;
		case "0_);[Red]\\(0\\)":
			//整数
			customProp.setFormat("number");
			content.setDisplayTexts(text);
			customProp.setDecimal(0);
			customProp.setThousands(false);
			if(CELLTYPE.STRING == excelCell.getType()){
				customProp.setIsValid(false);
			}
			break;	
		case "#,##0_);\\(#,##0\\)":
			//整数千分位
			setNumber(excelCell, 0, true);
			content.setDisplayTexts(excelCell.getShowText());
			customProp.setFormat("number");
			customProp.setThousands(true);
			customProp.setDecimal(0);
			if(CELLTYPE.STRING == excelCell.getType()){
				customProp.setIsValid(false);
			}
			break;
		case "#,##0_);[Red]\\(#,##0\\)":
			//整数千分位
			setNumber(excelCell, 0, true);
			content.setDisplayTexts(excelCell.getShowText());
			customProp.setFormat("number");
			customProp.setThousands(true);
			customProp.setDecimal(0);
			if(CELLTYPE.STRING == excelCell.getType()){
				customProp.setIsValid(false);
			}
			break;	
		case "\"¥\"#,##0_);\\(\"¥\"#,##0\\)":
			//货币
			setCurrency(excelCell, 0, "¥");
			customProp.setFormat("currency");
			customProp.setCurrencySign("¥");
			customProp.setDecimal(0);
			customProp.setThousands(true);
			content.setDisplayTexts(excelCell.getShowText());
			if(CELLTYPE.STRING == excelCell.getType()){
				customProp.setIsValid(false);
			}
			break;
		case "\"$\"#,##0_);\\(\"¥\"#,##0\\)":
			//货币
			setCurrency(excelCell, 0, "$");
			customProp.setFormat("currency");
			customProp.setDecimal(0);
			customProp.setCurrencySign("$");
			customProp.setThousands(true);
			content.setDisplayTexts(excelCell.getShowText());
			if(CELLTYPE.STRING == excelCell.getType()){
				customProp.setIsValid(false);
			}
			break;
		case "0%":
			setPercent(excelCell, 0);
			customProp.setFormat("percent");
			customProp.setDecimal(0);
			content.setDisplayTexts(excelCell.getShowText());
			if(CELLTYPE.STRING == excelCell.getType()){
				customProp.setIsValid(false);
			}
			break;
		default:
			if (dataFormate.startsWith("0.0") && dataFormate.endsWith("0\\)") && !dataFormate.contains(";[Red]")) {
				// 小数
				int index = dataFormate.indexOf("_");
				int decimal = index - 2;
				setNumber(excelCell, index - 2, false);
				customProp.setFormat("number");
				customProp.setDecimal(decimal);
				customProp.setThousands(false);
				content.setDisplayTexts(excelCell.getShowText());
				if(CELLTYPE.STRING == excelCell.getType()){
					customProp.setIsValid(false);
				}
			}else if (dataFormate.startsWith("0.0") && dataFormate.endsWith("0\\)") && dataFormate.contains(";[Red]")) {
				// 小数
				int index = dataFormate.indexOf("_");
				int decimal = index - 2;
				setNumber(excelCell, index - 2, false);
				customProp.setFormat("number");
				customProp.setDecimal(decimal);
				customProp.setThousands(false);
				String showText = excelCell.getShowText();
				if(showText != null && showText.contains("-")){
					showText = showText.replace("-", "");
					content.setDisplayTexts("("+showText+")");
					ExcelColor excelColor  = excelCell.getCellstyle().getFont().getColor();
					//excelColor.setAuto(false);
					excelColor.setRGB(255, 0, 0);
				}else{
					content.setDisplayTexts(showText);
				}
				
				if(CELLTYPE.STRING == excelCell.getType()){
					customProp.setIsValid(false);
				}
			}else if (dataFormate.startsWith("0.0") && !dataFormate.contains(";[Red]") && dataFormate.indexOf("_") != -1) {
				// 小数
				int index = dataFormate.indexOf("_");
				int decimal = index - 2;
				setNumber(excelCell, index - 2, false);
				customProp.setFormat("number");
				customProp.setDecimal(decimal);
				customProp.setThousands(false);
				content.setDisplayTexts(excelCell.getShowText());
				if(CELLTYPE.STRING == excelCell.getType()){
					customProp.setIsValid(false);
				}
			}else if (dataFormate.startsWith("0.0") && dataFormate.contains(";[Red]")) {
				// 小数
				int index = dataFormate.indexOf(";");
				int decimal = index - 2;
				setNumber(excelCell, index - 2, false);
				customProp.setFormat("number");
				customProp.setDecimal(decimal);
				customProp.setThousands(false);
				content.setDisplayTexts(excelCell.getShowText());
				if(CELLTYPE.STRING == excelCell.getType()){
					customProp.setIsValid(false);
				}
			}else if (dataFormate.startsWith("#,##0.0")) {
				// 小数千分位
				int index = dataFormate.indexOf("_");
				int decimal = index - 6;
				setNumber(excelCell, decimal, true);
				content.setDisplayTexts(excelCell.getShowText());
				customProp.setFormat("number");
				customProp.setDecimal(decimal);
				customProp.setThousands(true);
				if(CELLTYPE.STRING == excelCell.getType()){
					customProp.setIsValid(false);
				}
			} else if (dataFormate.startsWith("\"¥\"#,##0.0")) {
				//千分位货币
				int index = dataFormate.indexOf("_");
				int decimal = index - 9;
				setCurrency(excelCell, decimal, "¥");
				content.setDisplayTexts(excelCell.getShowText());
				customProp.setFormat("currency");
				customProp.setCurrencySign("¥");
				customProp.setDecimal(decimal);
				if(CELLTYPE.STRING == excelCell.getType()){
					customProp.setIsValid(false);
				}
			} else if (dataFormate.startsWith("\"$\"#,##0.0")) {
				//千分位货币
				int index = dataFormate.indexOf("_");
				int decimal = index - 9;
				setCurrency(excelCell, index - 9, "$");
				content.setDisplayTexts(excelCell.getShowText());
				customProp.setFormat("currency");
				customProp.setCurrencySign("$");
				customProp.setDecimal(decimal);
				if(CELLTYPE.STRING == excelCell.getType()){
					customProp.setIsValid(false);
				}
			} else if(dataFormate.startsWith("0.0") && dataFormate.endsWith("0%")){
				//小数百分比
				int index = dataFormate.indexOf("%");
				int decimal = index - 2;
				setPercent(excelCell, index-2);
				customProp.setDecimal(decimal);
				content.setDisplayTexts(excelCell.getShowText());
				customProp.setFormat("percent");
				if(CELLTYPE.STRING == excelCell.getType()){
					customProp.setIsValid(false);
				}
			}else{
				customProp.setFormat("normal");
				if(CELLTYPE.NUMERIC == excelCell.getType()){
					int d = ExcelFormat.getDecimalFormatDotcount(excelCell.getCellstyle().getDataformat());
					customProp.setDecimal(d);
				}
//				String displayText = ExcelFormat.getShowText(excelCell);
//				content.setDisplayTexts(displayText);
			}
			break;
		}
	}
	private static String getPrettyNumber(String number) {  
	    return BigDecimal.valueOf(Double.parseDouble(number))  
	            .stripTrailingZeros().toPlainString();  
	}  
	  
	public static void main(String[] args) {  
	    String intNumber = "12340.00";  
	    System.out.println(getPrettyNumber(intNumber));  
	    String doubleNumber = "00.340";  
	    System.out.println(getPrettyNumber(doubleNumber));  
	      
	    String eNumber = "1.2e3";  
	    System.out.println(getPrettyNumber(eNumber));  
	} 
}
