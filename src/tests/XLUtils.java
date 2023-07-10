package tests;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;


import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;



public class XLUtils {

	public static FileInputStream fi;
	public static FileOutputStream fo;
	public static XSSFWorkbook workbook;
	public static XSSFSheet sheet;
	public static XSSFRow row;
	public static XSSFCell cell;
    
    public static WebDriver driver;
	
	//to specify working File and Sheet
		public static void setExcelFile(String xlfile,String xlsheet) throws Exception {
			 
				try {
				
	   			fi = new FileInputStream(xlfile);
	   			workbook = new XSSFWorkbook(fi);
	   			sheet = workbook.getSheet(xlsheet);
				} catch (Exception e){
					throw (e);
				}
		}
	
		//To find number of Rows  data availability in a specified sheet
		
		public static int getRowCount(String xlfile,String xlsheet) throws IOException 
		{
			fi=new FileInputStream(xlfile);
			workbook=new XSSFWorkbook(fi);
			sheet=workbook.getSheet(xlsheet);
			int rowcount=sheet.getLastRowNum();
			workbook.close();
			fi.close();
			return rowcount;		
		}
		
		//to find number cells in a specified Row
		public static int getCellCount(String xlfile,String xlsheet,int rownum) throws IOException
		{
			fi=new FileInputStream(xlfile);
			workbook=new XSSFWorkbook(fi);
			sheet=workbook.getSheet(xlsheet);
			row=sheet.getRow(rownum);
			int cellcount=row.getLastCellNum();
			workbook.close();
			fi.close();
			return cellcount;
		}
		
		
		//to read cell value
		public static String getCellData(String xlfile,String xlsheet,int rownum,int colnum) throws IOException
		{
			fi=new FileInputStream(xlfile);
			workbook=new XSSFWorkbook(fi);
			sheet=workbook.getSheet(xlsheet);
			row=sheet.getRow(rownum);
			cell=row.getCell(colnum);
			String data;
			try 
			{
				DataFormatter formatter = new DataFormatter();
	            String cellData = formatter.formatCellValue(cell);
	            return cellData;
			}
			catch (Exception e) 
			{
				data="";
			}
			workbook.close();
			fi.close();
			return data;
		}
		
		//to set value
		public static void setCellData(String xlfile,String xlsheet,int rownum,int colnum,String data) throws IOException
		{
			fi=new FileInputStream(xlfile);
			workbook=new XSSFWorkbook(fi);
			sheet=workbook.getSheet(xlsheet);
			row=sheet.getRow(rownum);
			cell=row.createCell(colnum);
			cell.setCellValue(data);
			fo=new FileOutputStream(xlfile);
			workbook.write(fo);		
			workbook.close();
			fi.close();
			fo.close();
		}
	
	}


