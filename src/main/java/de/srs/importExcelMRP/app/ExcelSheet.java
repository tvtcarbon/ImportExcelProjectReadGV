package de.srs.importExcelMRP.app;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelSheet {
	
	private XSSFWorkbook workbook;
	private XSSFSheet sheet;
	
	public XSSFWorkbook getWorkbook() {
		return workbook;
	}
	public void setWorkbook(XSSFWorkbook workbook) {
		this.workbook = workbook;
	}
	
	public XSSFSheet getSheet() {
		return sheet;
	}

	public void setSheet(XSSFSheet sheet) {
		this.sheet = sheet;
	}
	
	public ExcelSheet() {
		super();
	}
	
	public ExcelSheet(XSSFWorkbook workbook, XSSFSheet sheet) {
		super();
		this.workbook = workbook;
		this.setSheet(sheet);
	}
	

}
