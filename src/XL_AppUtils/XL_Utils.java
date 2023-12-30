package XL_AppUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xslf.usermodel.XSLFSlideShowFactory;
import org.apache.poi.xssf.usermodel.IndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XL_Utils 
{

	public static 	FileInputStream fi;
	public static  Workbook wb;
	public static  Sheet ws;
	public static   Row row;
	public static  Cell cell;

	public static int getRowcount(String xlfile, String xlsheet) throws IOException 
	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws =wb.getSheet(xlsheet);
		int rowcount =    ws.getLastRowNum();	
		wb.close();
		return rowcount;

	}


	public static  short getcoloumncount(String xlfile, String xlsheet, int rownum) throws IOException 
	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws =	wb.getSheet(xlsheet);
		row =	ws.getRow(rownum);

		int colcount =	row.getLastCellNum();
		wb.close();
		return (short) colcount;

	}



	public static String getStringData(String xlfile , String xlsheet, int rownum ,int cellnum) throws IOException 
	{

		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row =ws.getRow(rownum);
		cell = row.getCell(cellnum);
		String empname1 = cell.getStringCellValue();
		wb.close();
		return empname1;				

	}


	public static double getNumericData(String xlfile, String xlsheet, int rownum ,int cellnum) throws IOException 

	{

		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws =  wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		cell = row.getCell(cellnum);

		double  salary  =	cell.getNumericCellValue();

		wb.close();
		return salary;


	}


	public static  boolean getBooleanData(String xlfile, String xlsheet, int rownum ,int cellnum) throws IOException 

	{

		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws =  wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		cell = row.getCell(cellnum);

		boolean status =	cell.getBooleanCellValue();

		wb.close();
		return status;

	}




	public static void getSetData(String xlfile, String xlsheet, int rownum ,int cellnum ,String data) throws IOException 

	{

		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws =  wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
	//	cell = row.getCell(cellnum);

		Cell cell =   row.createCell(cellnum);

		cell.setCellValue(data);

		FileOutputStream fo = new FileOutputStream(xlfile);
		wb.write(fo);

		wb.close();



	}	




	public static void FillGreencolour(String xlfile, String xlsheet, int rownum ,int cellnum ,String data) throws IOException 

	{

		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws =  wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		//cell = row.getCell(cellnum);

		Cell cell1 =   row.createCell(cellnum);
		//Cell cell2 = row.createCell(cellnum);
		
		cell1.setCellValue(data);
		//cell2.setCellValue(data);



		CellStyle passstyle =  wb.createCellStyle();
		passstyle.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
		passstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell1.setCellStyle(passstyle);  
             
		
		//CellStyle failstyle = wb.createCellStyle();
	//	failstyle.setFillForegroundColor(IndexedColors.RED.getIndex());
	//	failstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		//cell2.setCellStyle(failstyle);
		
		
		
		
		


		FileOutputStream fo = new FileOutputStream(xlfile);
		wb.write(fo);

		wb.close();
	}

	
	
	
	
	public static void FillRedcolour(String xlfile, String xlsheet, int rownum ,int cellnum ,String data) throws IOException 

	{

		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws =  wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		//cell = row.getCell(cellnum);

		Cell cell1 =   row.createCell(cellnum);
	
		cell1.setCellValue(data);
	
	
		CellStyle failstyle = wb.createCellStyle();
		failstyle.setFillForegroundColor(IndexedColors.RED.getIndex());
		failstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell1.setCellStyle(failstyle);
		
	
	
		FileOutputStream fo = new FileOutputStream(xlfile);
		wb.write(fo);

		wb.close();
	
	
	
	
	}

}
