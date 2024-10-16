package utilities;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileUtile 
{
	XSSFWorkbook wb;
	
	//Write Constructor for reading Excel path
	public ExcelFileUtile(String ExcelPath) throws Throwable
	{
		FileInputStream fi = new FileInputStream(ExcelPath);
		wb = new XSSFWorkbook(fi);

	}
	
	// Method for Counting No of rows in a Sheet
	public int rowCount(String sheetName)
	{
		
		return wb.getSheet(sheetName).getLastRowNum();

	}
	
	// Method For Reading Cell Data
	public String getCellData(String sheetName, int row, int column)
	{
		
		
		String data="";
		if(wb.getSheet(sheetName).getRow(row).getCell(column).getCellType()==CellType.NUMERIC)
		{
			int celldata =(int) wb.getSheet(sheetName).getRow(row).getCell(column).getNumericCellValue();
			data =String.valueOf(celldata);
		}
		else
		{
			data =wb.getSheet(sheetName).getRow(row).getCell(column).getStringCellValue();
		}
		return data;

	}
	
	// Method for Writing results into new WB
	public void setCellData(String sheetName, int row, int column, String status, String WriteExcel ) throws Throwable
	{
		//get sheet from wb
		XSSFSheet ws = wb.getSheet(sheetName);
		
		//get row from sheet
		XSSFRow rowNum =ws.getRow(row);
		
		//create cell in row
		XSSFCell cell =rowNum.createCell(column);
		
		cell.setCellValue(status);
		if(status.equalsIgnoreCase("Pass"))
		{
			XSSFCellStyle style = wb.createCellStyle();
			XSSFFont font = wb.createFont();
			font.setColor(IndexedColors.GREEN.getIndex());
			font.setBold(true);
			style.setFont(font);
			rowNum.getCell(column).setCellStyle(style);
		}
		
		else if(status.equalsIgnoreCase("Fail"))
		{
			XSSFCellStyle style = wb.createCellStyle();
			XSSFFont font = wb.createFont();
			font.setColor(IndexedColors.RED.getIndex());
			font.setBold(true);
			style.setFont(font);
			rowNum.getCell(column).setCellStyle(style);
		}
		
		else if(status.equalsIgnoreCase("Blocked"))
		{
			XSSFCellStyle style = wb.createCellStyle();
			XSSFFont font = wb.createFont();
			font.setColor(IndexedColors.BLUE.getIndex());
			font.setBold(true);
			style.setFont(font);
			rowNum.getCell(column).setCellStyle(style);
		}
		FileOutputStream fo = new FileOutputStream(WriteExcel);
		wb.write(fo);

	}	

}
