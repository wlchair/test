package com.acmr.excel.model.complete;

import java.io.Serializable;

public class CustomProp implements Serializable {
	private String background = "rgb(255,255,255)";
	private String bgRgbColor;
	private String format = "normal";
	private String remarket;
	private Integer decimal;
	private Boolean thousands;
	private String dateFormat;
	private String currencySign;
	private String comment;
	/**
	 * 文本内容，与设置类型是否匹配
	 */
	private Boolean isValid = true;

	public String getBackground() {
		return background;
	}

	public void setBackground(String background) {
		this.background = background;
	}

	public String getBgRgbColor() {
		return bgRgbColor;
	}

	public void setBgRgbColor(String bgRgbColor) {
		this.bgRgbColor = bgRgbColor;
	}

	public String getFormat() {
		return format;
	}

	public void setFormat(String format) {
		this.format = format;
	}

	public String getRemarket() {
		return remarket;
	}

	public void setRemarket(String remarket) {
		this.remarket = remarket;
	}

	public Integer getDecimal() {
		return decimal;
	}

	public void setDecimal(Integer decimal) {
		this.decimal = decimal;
	}

	public Boolean getThousands() {
		return thousands;
	}

	public void setThousands(Boolean thousands) {
		this.thousands = thousands;
	}

	public String getDateFormat() {
		return dateFormat;
	}

	public void setDateFormat(String dateFormat) {
		this.dateFormat = dateFormat;
	}

	public String getCurrencySign() {
		return currencySign;
	}

	public void setCurrencySign(String currencySign) {
		this.currencySign = currencySign;
	}

	public String getComment() {
		return comment;
	}

	public void setComment(String comment) {
		this.comment = comment;
	}

	public Boolean getIsValid() {
		return isValid;
	}

	public void setIsValid(Boolean isValid) {
		this.isValid = isValid;
	}

}
