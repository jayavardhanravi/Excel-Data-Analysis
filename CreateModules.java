
package maid.myapi.com;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Iterator;
import java.util.Random;

import maid.mycontent.orig.DataBean;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;

public class CreateModules {

	public String createBlankExcelSheet(String filePath){
		System.out.println("Started creating the output excel file");
		Workbook wb = new XSSFWorkbook();
		Random rn = new Random();
		int rnInt = rn.nextInt();
		String fileName= "OutputExcelFile"+String.valueOf(rnInt)+"jaya.xlsx";
	      FileOutputStream out = null;
		try {
			out = new FileOutputStream(fileName);
			wb.write(out);
			out.close();
			wb.close();
			System.out.println("Created the out excel file, with Name :"+fileName);
			return fileName;
		} catch (Exception e) {
			e.printStackTrace();
			return fileName;
		}
	
	}
	
	public void createSheetsInExcel(Workbook wb){
		Sheet sheet1 = wb.createSheet("JAN");
		Sheet sheet2 = wb.createSheet("FEB");
		Sheet sheet3 = wb.createSheet("MAR");
		Sheet sheet4 = wb.createSheet("APR");
		Sheet sheet5 = wb.createSheet("MAY");
		Sheet sheet6 = wb.createSheet("JUN");
		Sheet sheet7 = wb.createSheet("JUL");
		Sheet sheet8 = wb.createSheet("AUG");
		Sheet sheet9 = wb.createSheet("SEP");
		Sheet sheet10 = wb.createSheet("OCT");
		Sheet sheet11 = wb.createSheet("NOV");
		Sheet sheet12 = wb.createSheet("DEC");
		System.out.println("Generated the sheets in excel file");
	}
	
	public void addDataToSheetJan(Workbook wb, Workbook originalWb){
		CellStyle style = wb.createCellStyle();
		style = wb.createCellStyle();
	    style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    style.setBorderBottom(CellStyle.BORDER_THICK);
	    style.setBorderTop(CellStyle.BORDER_THICK);
	    style.setBorderLeft(CellStyle.BORDER_THICK);
	    style.setBorderRight(CellStyle.BORDER_THICK);
		Row row = wb.getSheet("JAN").createRow((short)0);
		Cell cell = row.createCell(0);
		cell.setCellValue("Item (SKU)");
		wb.getSheet("JAN").autoSizeColumn(0);
		cell.setCellStyle(style);
		Cell cell1 = row.createCell(1);
		cell1.setCellValue("Description");
		wb.getSheet("JAN").autoSizeColumn(1);
		wb.getSheet("JAN").setColumnWidth(1, 50*256);
		cell1.setCellStyle(style);
		
		Calendar cal = Calendar.getInstance();
		cal.set(Calendar.YEAR, 2014);
	    cal.set(Calendar.MONTH, 0);
	    cal.set(Calendar.DAY_OF_MONTH, 1);
	    int maxDay = cal.getActualMaximum(Calendar.DAY_OF_MONTH);
	    SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
	    for (int i = 1; i <= maxDay; i++) {
	        Cell ceDate = row.createCell(i+1);
	    	ceDate.setCellValue(df.format(cal.getTime()));
			wb.getSheet("JAN").autoSizeColumn(i+1);
			ceDate.setCellStyle(style);
			cal.set(Calendar.DAY_OF_MONTH, i + 1);
	    }
		
		ArrayList<DataBean> arMain = new ArrayList<DataBean>();
		for(int i=2; i< originalWb.getSheet("JAN").getPhysicalNumberOfRows(); i++){
			ArrayList<DataBean> arSub1 = new ArrayList<DataBean>();
			if(!originalWb.getSheet("JAN").getRow(i).getCell(0).getStringCellValue().equals(null)){
				DataBean dbeanObj = new DataBean();
				dbeanObj.setItemNumber(String.valueOf(originalWb.getSheet("JAN").getRow(i).getCell(0).getStringCellValue()));
				dbeanObj.setDescription(String.valueOf(originalWb.getSheet("JAN").getRow(i).getCell(1).getStringCellValue()));
				arSub1.add(dbeanObj);
			}
			arMain.addAll(arSub1);
		}
		DataBean obj = null;
		
		for(int i=1;i<arMain.size();i++){
			obj = (DataBean) arMain.get(i);
			
			if(!obj.getItemNumber().equals(null)){
				Row r = wb.getSheet("JAN").createRow((short)i);
				Cell ce = r.createCell(0);
				ce.setCellValue(Integer.valueOf(obj.getItemNumber().trim()));
				Cell ce1 = r.createCell(1);
				ce1.setCellValue(obj.getDescription());
			}
			
		}

	}
	
	
	public void addDataToSheetFeb(Workbook wb, Workbook originalWb){
		CellStyle style = wb.createCellStyle();
		style = wb.createCellStyle();
	    style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    style.setBorderBottom(CellStyle.BORDER_THICK);
	    style.setBorderTop(CellStyle.BORDER_THICK);
	    style.setBorderLeft(CellStyle.BORDER_THICK);
	    style.setBorderRight(CellStyle.BORDER_THICK);
		Row row = wb.getSheet("FEB").createRow((short)0);
		Cell cell = row.createCell(0);
		cell.setCellValue("Item (SKU)");
		wb.getSheet("FEB").autoSizeColumn(0);
		cell.setCellStyle(style);
		Cell cell1 = row.createCell(1);
		cell1.setCellValue("Description");
		wb.getSheet("FEB").autoSizeColumn(1);
		wb.getSheet("FEB").setColumnWidth(1, 50*256);
		cell1.setCellStyle(style);
		
		Calendar cal = Calendar.getInstance();
		cal.set(Calendar.YEAR, 2014);
	    cal.set(Calendar.MONTH, 1);
	    cal.set(Calendar.DAY_OF_MONTH, 1);
	    int maxDay = cal.getActualMaximum(Calendar.DAY_OF_MONTH);
	    SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
	    for (int i = 1; i <= maxDay; i++) {
	        Cell ceDate = row.createCell(i+1);
	    	ceDate.setCellValue(df.format(cal.getTime()));
			wb.getSheet("FEB").autoSizeColumn(i+1);
			ceDate.setCellStyle(style);
			cal.set(Calendar.DAY_OF_MONTH, i + 1);
	    }
		
		ArrayList<DataBean> arMain = new ArrayList<DataBean>();
		for(int i=2; i< originalWb.getSheet("FEB").getPhysicalNumberOfRows(); i++){
			ArrayList<DataBean> arSub1 = new ArrayList<DataBean>();
			DataBean dbeanObj = new DataBean();
			dbeanObj.setItemNumber(String.valueOf(originalWb.getSheet("FEB").getRow(i).getCell(0).getStringCellValue()));
			dbeanObj.setDescription(String.valueOf(originalWb.getSheet("FEB").getRow(i).getCell(1).getStringCellValue()));
			arSub1.add(dbeanObj);
			arMain.addAll(arSub1);
		}
		DataBean obj = null;
		
		for(int i=1;i<arMain.size();i++){
			obj = (DataBean) arMain.get(i);
			
			if(!obj.getItemNumber().equals(null)){
			Row r = wb.getSheet("FEB").createRow((short)i);
			Cell ce = r.createCell(0);
			ce.setCellValue(Integer.valueOf(obj.getItemNumber().trim()));
			Cell ce1 = r.createCell(1);
			ce1.setCellValue(obj.getDescription());
			}
		}

	}
	
	
	public void addDataToSheetMar(Workbook wb, Workbook originalWb){
		CellStyle style = wb.createCellStyle();
		style = wb.createCellStyle();
	    style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    style.setBorderBottom(CellStyle.BORDER_THICK);
	    style.setBorderTop(CellStyle.BORDER_THICK);
	    style.setBorderLeft(CellStyle.BORDER_THICK);
	    style.setBorderRight(CellStyle.BORDER_THICK);
		Row row = wb.getSheet("MAR").createRow((short)0);
		Cell cell = row.createCell(0);
		cell.setCellValue("Item (SKU)");
		wb.getSheet("MAR").autoSizeColumn(0);
		cell.setCellStyle(style);
		Cell cell1 = row.createCell(1);
		cell1.setCellValue("Description");
		wb.getSheet("MAR").autoSizeColumn(1);
		wb.getSheet("MAR").setColumnWidth(1, 50*256);
		cell1.setCellStyle(style);
		
		Calendar cal = Calendar.getInstance();
		cal.set(Calendar.YEAR, 2014);
	    cal.set(Calendar.MONTH, 2);
	    cal.set(Calendar.DAY_OF_MONTH, 1);
	    int maxDay = cal.getActualMaximum(Calendar.DAY_OF_MONTH);
	    SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
	    for (int i = 1; i <= maxDay; i++) {
	        Cell ceDate = row.createCell(i+1);
	    	ceDate.setCellValue(df.format(cal.getTime()));
			wb.getSheet("MAR").autoSizeColumn(i+1);
			ceDate.setCellStyle(style);
			cal.set(Calendar.DAY_OF_MONTH, i + 1);
	    }
		
		ArrayList<DataBean> arMain = new ArrayList<DataBean>();
		for(int i=2; i< originalWb.getSheet("MAR").getPhysicalNumberOfRows(); i++){
			ArrayList<DataBean> arSub1 = new ArrayList<DataBean>();
			DataBean dbeanObj = new DataBean();
			dbeanObj.setItemNumber(String.valueOf(originalWb.getSheet("MAR").getRow(i).getCell(0).getStringCellValue()));
			dbeanObj.setDescription(String.valueOf(originalWb.getSheet("MAR").getRow(i).getCell(1).getStringCellValue()));
			arSub1.add(dbeanObj);
			arMain.addAll(arSub1);
		}
		DataBean obj = null;
		
		for(int i=1;i<arMain.size();i++){
			obj = (DataBean) arMain.get(i);
			
			if(!obj.getItemNumber().equals(null)){
			Row r = wb.getSheet("MAR").createRow((short)i);
			Cell ce = r.createCell(0);
			ce.setCellValue(Integer.valueOf(obj.getItemNumber().trim()));
			Cell ce1 = r.createCell(1);
			ce1.setCellValue(obj.getDescription());
			}
		}

	}
	
	public void addDataToSheetApr(Workbook wb, Workbook originalWb){
		CellStyle style = wb.createCellStyle();
		style = wb.createCellStyle();
	    style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    style.setBorderBottom(CellStyle.BORDER_THICK);
	    style.setBorderTop(CellStyle.BORDER_THICK);
	    style.setBorderLeft(CellStyle.BORDER_THICK);
	    style.setBorderRight(CellStyle.BORDER_THICK);
		Row row = wb.getSheet("APR").createRow((short)0);
		Cell cell = row.createCell(0);
		cell.setCellValue("Item (SKU)");
		wb.getSheet("APR").autoSizeColumn(0);
		cell.setCellStyle(style);
		Cell cell1 = row.createCell(1);
		cell1.setCellValue("Description");
		wb.getSheet("APR").autoSizeColumn(1);
		wb.getSheet("APR").setColumnWidth(1, 50*256);
		cell1.setCellStyle(style);
		
		Calendar cal = Calendar.getInstance();
		cal.set(Calendar.YEAR, 2014);
	    cal.set(Calendar.MONTH, 3);
	    cal.set(Calendar.DAY_OF_MONTH, 1);
	    int maxDay = cal.getActualMaximum(Calendar.DAY_OF_MONTH);
	    SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
	    for (int i = 1; i <= maxDay; i++) {
	        Cell ceDate = row.createCell(i+1);
	    	ceDate.setCellValue(df.format(cal.getTime()));
			wb.getSheet("APR").autoSizeColumn(i+1);
			ceDate.setCellStyle(style);
			cal.set(Calendar.DAY_OF_MONTH, i + 1);
	    }
		
		ArrayList<DataBean> arMain = new ArrayList<DataBean>();
		for(int i=2; i< originalWb.getSheet("APR").getPhysicalNumberOfRows(); i++){
			ArrayList<DataBean> arSub1 = new ArrayList<DataBean>();
			DataBean dbeanObj = new DataBean();
			dbeanObj.setItemNumber(String.valueOf(originalWb.getSheet("APR").getRow(i).getCell(0).getStringCellValue()));
			dbeanObj.setDescription(String.valueOf(originalWb.getSheet("APR").getRow(i).getCell(1).getStringCellValue()));
			arSub1.add(dbeanObj);
			arMain.addAll(arSub1);
		}
		DataBean obj = null;
		
		for(int i=1;i<arMain.size();i++){
			obj = (DataBean) arMain.get(i);
			
			if(!obj.getItemNumber().equals(null)){
			Row r = wb.getSheet("APR").createRow((short)i);
			Cell ce = r.createCell(0);
			ce.setCellValue(Integer.valueOf(obj.getItemNumber().trim()));
			Cell ce1 = r.createCell(1);
			ce1.setCellValue(obj.getDescription());
			}
		}

	}
	
	
	public void addDataToSheetMay(Workbook wb, Workbook originalWb){
		CellStyle style = wb.createCellStyle();
		style = wb.createCellStyle();
	    style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    style.setBorderBottom(CellStyle.BORDER_THICK);
	    style.setBorderTop(CellStyle.BORDER_THICK);
	    style.setBorderLeft(CellStyle.BORDER_THICK);
	    style.setBorderRight(CellStyle.BORDER_THICK);
		Row row = wb.getSheet("MAY").createRow((short)0);
		Cell cell = row.createCell(0);
		cell.setCellValue("Item (SKU)");
		wb.getSheet("MAY").autoSizeColumn(0);
		cell.setCellStyle(style);
		Cell cell1 = row.createCell(1);
		cell1.setCellValue("Description");
		wb.getSheet("MAY").autoSizeColumn(1);
		wb.getSheet("MAY").setColumnWidth(1, 50*256);
		cell1.setCellStyle(style);
		
		Calendar cal = Calendar.getInstance();
		cal.set(Calendar.YEAR, 2014);
	    cal.set(Calendar.MONTH, 4);
	    cal.set(Calendar.DAY_OF_MONTH, 1);
	    int maxDay = cal.getActualMaximum(Calendar.DAY_OF_MONTH);
	    SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
	    for (int i = 1; i <= maxDay; i++) {
	        Cell ceDate = row.createCell(i+1);
	    	ceDate.setCellValue(df.format(cal.getTime()));
			wb.getSheet("MAY").autoSizeColumn(i+1);
			ceDate.setCellStyle(style);
			cal.set(Calendar.DAY_OF_MONTH, i + 1);
	    }
		
		ArrayList<DataBean> arMain = new ArrayList<DataBean>();
		for(int i=2; i< originalWb.getSheet("MAY").getPhysicalNumberOfRows(); i++){
			ArrayList<DataBean> arSub1 = new ArrayList<DataBean>();
			DataBean dbeanObj = new DataBean();
			dbeanObj.setItemNumber(String.valueOf(originalWb.getSheet("MAY").getRow(i).getCell(0).getStringCellValue()));
			dbeanObj.setDescription(String.valueOf(originalWb.getSheet("MAY").getRow(i).getCell(1).getStringCellValue()));
			arSub1.add(dbeanObj);
			arMain.addAll(arSub1);
		}
		DataBean obj = null;
		
		for(int i=1;i<arMain.size();i++){
			obj = (DataBean) arMain.get(i);
			
			if(!obj.getItemNumber().equals(null)){
			Row r = wb.getSheet("MAY").createRow((short)i);
			Cell ce = r.createCell(0);
			ce.setCellValue(Integer.valueOf(obj.getItemNumber().trim()));
			Cell ce1 = r.createCell(1);
			ce1.setCellValue(obj.getDescription());
			}
		}

	}
	
	public void addDataToSheetJun(Workbook wb, Workbook originalWb){
		CellStyle style = wb.createCellStyle();
		style = wb.createCellStyle();
	    style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    style.setBorderBottom(CellStyle.BORDER_THICK);
	    style.setBorderTop(CellStyle.BORDER_THICK);
	    style.setBorderLeft(CellStyle.BORDER_THICK);
	    style.setBorderRight(CellStyle.BORDER_THICK);
		Row row = wb.getSheet("JUN").createRow((short)0);
		Cell cell = row.createCell(0);
		cell.setCellValue("Item (SKU)");
		wb.getSheet("JUN").autoSizeColumn(0);
		cell.setCellStyle(style);
		Cell cell1 = row.createCell(1);
		cell1.setCellValue("Description");
		wb.getSheet("JUN").autoSizeColumn(1);
		wb.getSheet("JUN").setColumnWidth(1, 50*256);
		cell1.setCellStyle(style);
		
		Calendar cal = Calendar.getInstance();
		cal.set(Calendar.YEAR, 2014);
	    cal.set(Calendar.MONTH, 5);
	    cal.set(Calendar.DAY_OF_MONTH, 1);
	    int maxDay = cal.getActualMaximum(Calendar.DAY_OF_MONTH);
	    SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
	    for (int i = 1; i <= maxDay; i++) {
	        Cell ceDate = row.createCell(i+1);
	    	ceDate.setCellValue(df.format(cal.getTime()));
			wb.getSheet("JUN").autoSizeColumn(i+1);
			ceDate.setCellStyle(style);
			cal.set(Calendar.DAY_OF_MONTH, i + 1);
	    }
		
		ArrayList<DataBean> arMain = new ArrayList<DataBean>();
		for(int i=2; i< originalWb.getSheet("JUN").getPhysicalNumberOfRows(); i++){
			ArrayList<DataBean> arSub1 = new ArrayList<DataBean>();
			DataBean dbeanObj = new DataBean();
			dbeanObj.setItemNumber(String.valueOf(originalWb.getSheet("JUN").getRow(i).getCell(0).getStringCellValue()));
			dbeanObj.setDescription(String.valueOf(originalWb.getSheet("JUN").getRow(i).getCell(1).getStringCellValue()));
			arSub1.add(dbeanObj);
			arMain.addAll(arSub1);
		}
		DataBean obj = null;
		
		for(int i=1;i<arMain.size();i++){
			obj = (DataBean) arMain.get(i);
			
			if(!obj.getItemNumber().equals(null)){
			Row r = wb.getSheet("JUN").createRow((short)i);
			Cell ce = r.createCell(0);
			ce.setCellValue(Integer.valueOf(obj.getItemNumber().trim()));
			Cell ce1 = r.createCell(1);
			ce1.setCellValue(obj.getDescription());
			}
		}

	}
	
	public void addDataToSheetJul(Workbook wb, Workbook originalWb){
		CellStyle style = wb.createCellStyle();
		style = wb.createCellStyle();
	    style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    style.setBorderBottom(CellStyle.BORDER_THICK);
	    style.setBorderTop(CellStyle.BORDER_THICK);
	    style.setBorderLeft(CellStyle.BORDER_THICK);
	    style.setBorderRight(CellStyle.BORDER_THICK);
		Row row = wb.getSheet("JUL").createRow((short)0);
		Cell cell = row.createCell(0);
		cell.setCellValue("Item (SKU)");
		wb.getSheet("JUL").autoSizeColumn(0);
		cell.setCellStyle(style);
		Cell cell1 = row.createCell(1);
		cell1.setCellValue("Description");
		wb.getSheet("JUL").autoSizeColumn(1);
		wb.getSheet("JUL").setColumnWidth(1, 50*256);
		cell1.setCellStyle(style);
		
		Calendar cal = Calendar.getInstance();
		cal.set(Calendar.YEAR, 2014);
	    cal.set(Calendar.MONTH, 6);
	    cal.set(Calendar.DAY_OF_MONTH, 1);
	    int maxDay = cal.getActualMaximum(Calendar.DAY_OF_MONTH);
	    SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
	    for (int i = 1; i <= maxDay; i++) {
	        Cell ceDate = row.createCell(i+1);
	    	ceDate.setCellValue(df.format(cal.getTime()));
			wb.getSheet("JUL").autoSizeColumn(i+1);
			ceDate.setCellStyle(style);
			cal.set(Calendar.DAY_OF_MONTH, i + 1);
	    }
		
		ArrayList<DataBean> arMain = new ArrayList<DataBean>();
		for(int i=2; i< originalWb.getSheet("JUL").getPhysicalNumberOfRows(); i++){
			ArrayList<DataBean> arSub1 = new ArrayList<DataBean>();
			DataBean dbeanObj = new DataBean();
			dbeanObj.setItemNumber(String.valueOf(originalWb.getSheet("JUL").getRow(i).getCell(0).getStringCellValue()));
			dbeanObj.setDescription(String.valueOf(originalWb.getSheet("JUL").getRow(i).getCell(1).getStringCellValue()));
			arSub1.add(dbeanObj);
			arMain.addAll(arSub1);
		}
		DataBean obj = null;
		
		for(int i=1;i<arMain.size();i++){
			obj = (DataBean) arMain.get(i);
			
			if(!obj.getItemNumber().equals(null)){
			Row r = wb.getSheet("JUL").createRow((short)i);
			Cell ce = r.createCell(0);
			ce.setCellValue(Integer.valueOf(obj.getItemNumber().trim()));
			Cell ce1 = r.createCell(1);
			ce1.setCellValue(obj.getDescription());
			}
		}

	}
	
	public void addDataToSheetAug(Workbook wb, Workbook originalWb){
		CellStyle style = wb.createCellStyle();
		style = wb.createCellStyle();
	    style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    style.setBorderBottom(CellStyle.BORDER_THICK);
	    style.setBorderTop(CellStyle.BORDER_THICK);
	    style.setBorderLeft(CellStyle.BORDER_THICK);
	    style.setBorderRight(CellStyle.BORDER_THICK);
		Row row = wb.getSheet("AUG").createRow((short)0);
		Cell cell = row.createCell(0);
		cell.setCellValue("Item (SKU)");
		wb.getSheet("AUG").autoSizeColumn(0);
		cell.setCellStyle(style);
		Cell cell1 = row.createCell(1);
		cell1.setCellValue("Description");
		wb.getSheet("AUG").autoSizeColumn(1);
		wb.getSheet("AUG").setColumnWidth(1, 50*256);
		cell1.setCellStyle(style);
		
		Calendar cal = Calendar.getInstance();
		cal.set(Calendar.YEAR, 2014);
	    cal.set(Calendar.MONTH, 7);
	    cal.set(Calendar.DAY_OF_MONTH, 1);
	    int maxDay = cal.getActualMaximum(Calendar.DAY_OF_MONTH);
	    SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
	    for (int i = 1; i <= maxDay; i++) {
	        Cell ceDate = row.createCell(i+1);
	    	ceDate.setCellValue(df.format(cal.getTime()));
			wb.getSheet("AUG").autoSizeColumn(i+1);
			ceDate.setCellStyle(style);
			cal.set(Calendar.DAY_OF_MONTH, i + 1);
	    }
		
		ArrayList<DataBean> arMain = new ArrayList<DataBean>();
		for(int i=2; i< originalWb.getSheet("AUG").getPhysicalNumberOfRows(); i++){
			ArrayList<DataBean> arSub1 = new ArrayList<DataBean>();
			DataBean dbeanObj = new DataBean();
			dbeanObj.setItemNumber(String.valueOf(originalWb.getSheet("AUG").getRow(i).getCell(0).getStringCellValue()));
			dbeanObj.setDescription(String.valueOf(originalWb.getSheet("AUG").getRow(i).getCell(1).getStringCellValue()));
			arSub1.add(dbeanObj);
			arMain.addAll(arSub1);
		}
		DataBean obj = null;
		
		for(int i=1;i<arMain.size();i++){
			obj = (DataBean) arMain.get(i);
			
			if(!obj.getItemNumber().equals(null)){
			Row r = wb.getSheet("AUG").createRow((short)i);
			Cell ce = r.createCell(0);
			ce.setCellValue(Integer.valueOf(obj.getItemNumber().trim()));
			Cell ce1 = r.createCell(1);
			ce1.setCellValue(obj.getDescription());
			}
		}

	}
	
	public void addDataToSheetSep(Workbook wb, Workbook originalWb){
		CellStyle style = wb.createCellStyle();
		style = wb.createCellStyle();
	    style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    style.setBorderBottom(CellStyle.BORDER_THICK);
	    style.setBorderTop(CellStyle.BORDER_THICK);
	    style.setBorderLeft(CellStyle.BORDER_THICK);
	    style.setBorderRight(CellStyle.BORDER_THICK);
		Row row = wb.getSheet("SEP").createRow((short)0);
		Cell cell = row.createCell(0);
		cell.setCellValue("Item (SKU)");
		wb.getSheet("SEP").autoSizeColumn(0);
		cell.setCellStyle(style);
		Cell cell1 = row.createCell(1);
		cell1.setCellValue("Description");
		wb.getSheet("SEP").autoSizeColumn(1);
		wb.getSheet("SEP").setColumnWidth(1, 50*256);
		cell1.setCellStyle(style);
		
		Calendar cal = Calendar.getInstance();
		cal.set(Calendar.YEAR, 2014);
	    cal.set(Calendar.MONTH, 8);
	    cal.set(Calendar.DAY_OF_MONTH, 1);
	    int maxDay = cal.getActualMaximum(Calendar.DAY_OF_MONTH);
	    SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
	    for (int i = 1; i <= maxDay; i++) {
	        Cell ceDate = row.createCell(i+1);
	    	ceDate.setCellValue(df.format(cal.getTime()));
			wb.getSheet("SEP").autoSizeColumn(i+1);
			ceDate.setCellStyle(style);
			cal.set(Calendar.DAY_OF_MONTH, i + 1);
	    }
		
		ArrayList<DataBean> arMain = new ArrayList<DataBean>();
		for(int i=2; i< originalWb.getSheet("SEP").getPhysicalNumberOfRows(); i++){
			ArrayList<DataBean> arSub1 = new ArrayList<DataBean>();
			DataBean dbeanObj = new DataBean();
			dbeanObj.setItemNumber(String.valueOf(originalWb.getSheet("SEP").getRow(i).getCell(0).getStringCellValue()));
			dbeanObj.setDescription(String.valueOf(originalWb.getSheet("SEP").getRow(i).getCell(1).getStringCellValue()));
			arSub1.add(dbeanObj);
			arMain.addAll(arSub1);
		}
		DataBean obj = null;
		
		for(int i=1;i<arMain.size();i++){
			obj = (DataBean) arMain.get(i);
			
			if(!obj.getItemNumber().equals(null)){
			Row r = wb.getSheet("SEP").createRow((short)i);
			Cell ce = r.createCell(0);
			ce.setCellValue(Integer.valueOf(obj.getItemNumber().trim()));
			Cell ce1 = r.createCell(1);
			ce1.setCellValue(obj.getDescription());
			}
		}

	}
	
	public void addDataToSheetOct(Workbook wb, Workbook originalWb){
		CellStyle style = wb.createCellStyle();
		style = wb.createCellStyle();
	    style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    style.setBorderBottom(CellStyle.BORDER_THICK);
	    style.setBorderTop(CellStyle.BORDER_THICK);
	    style.setBorderLeft(CellStyle.BORDER_THICK);
	    style.setBorderRight(CellStyle.BORDER_THICK);
		Row row = wb.getSheet("OCT").createRow((short)0);
		Cell cell = row.createCell(0);
		cell.setCellValue("Item (SKU)");
		wb.getSheet("OCT").autoSizeColumn(0);
		cell.setCellStyle(style);
		Cell cell1 = row.createCell(1);
		cell1.setCellValue("Description");
		wb.getSheet("OCT").autoSizeColumn(1);
		wb.getSheet("OCT").setColumnWidth(1, 50*256);
		cell1.setCellStyle(style);
		
		Calendar cal = Calendar.getInstance();
		cal.set(Calendar.YEAR, 2014);
	    cal.set(Calendar.MONTH, 9);
	    cal.set(Calendar.DAY_OF_MONTH, 1);
	    int maxDay = cal.getActualMaximum(Calendar.DAY_OF_MONTH);
	    SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
	    for (int i = 1; i <= maxDay; i++) {
	        Cell ceDate = row.createCell(i+1);
	    	ceDate.setCellValue(df.format(cal.getTime()));
			wb.getSheet("OCT").autoSizeColumn(i+1);
			ceDate.setCellStyle(style);
			cal.set(Calendar.DAY_OF_MONTH, i + 1);
	    }
		
		ArrayList<DataBean> arMain = new ArrayList<DataBean>();
		for(int i=2; i< originalWb.getSheet("OCT").getPhysicalNumberOfRows(); i++){
			ArrayList<DataBean> arSub1 = new ArrayList<DataBean>();
			DataBean dbeanObj = new DataBean();
			dbeanObj.setItemNumber(String.valueOf(originalWb.getSheet("OCT").getRow(i).getCell(0).getStringCellValue()));
			dbeanObj.setDescription(String.valueOf(originalWb.getSheet("OCT").getRow(i).getCell(1).getStringCellValue()));
			arSub1.add(dbeanObj);
			arMain.addAll(arSub1);
		}
		DataBean obj = null;
		
		for(int i=1;i<arMain.size();i++){
			obj = (DataBean) arMain.get(i);
			
			if(!obj.getItemNumber().equals(null)){
			Row r = wb.getSheet("OCT").createRow((short)i);
			Cell ce = r.createCell(0);
			ce.setCellValue(Integer.valueOf(obj.getItemNumber().trim()));
			Cell ce1 = r.createCell(1);
			ce1.setCellValue(obj.getDescription());
			}
		}

	}
	
	
	public void addDataToSheetNov(Workbook wb, Workbook originalWb){
		CellStyle style = wb.createCellStyle();
		style = wb.createCellStyle();
	    style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    style.setBorderBottom(CellStyle.BORDER_THICK);
	    style.setBorderTop(CellStyle.BORDER_THICK);
	    style.setBorderLeft(CellStyle.BORDER_THICK);
	    style.setBorderRight(CellStyle.BORDER_THICK);
		Row row = wb.getSheet("NOV").createRow((short)0);
		Cell cell = row.createCell(0);
		cell.setCellValue("Item (SKU)");
		wb.getSheet("NOV").autoSizeColumn(0);
		cell.setCellStyle(style);
		Cell cell1 = row.createCell(1);
		cell1.setCellValue("Description");
		wb.getSheet("NOV").autoSizeColumn(1);
		wb.getSheet("NOV").setColumnWidth(1, 50*256);
		cell1.setCellStyle(style);
		
		Calendar cal = Calendar.getInstance();
		cal.set(Calendar.YEAR, 2014);
	    cal.set(Calendar.MONTH, 10);
	    cal.set(Calendar.DAY_OF_MONTH, 1);
	    int maxDay = cal.getActualMaximum(Calendar.DAY_OF_MONTH);
	    SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
	    for (int i = 1; i <= maxDay; i++) {
	        Cell ceDate = row.createCell(i+1);
	    	ceDate.setCellValue(df.format(cal.getTime()));
			wb.getSheet("NOV").autoSizeColumn(i+1);
			ceDate.setCellStyle(style);
			cal.set(Calendar.DAY_OF_MONTH, i + 1);
	    }
		
		ArrayList<DataBean> arMain = new ArrayList<DataBean>();
		for(int i=2; i< originalWb.getSheet("NOV").getPhysicalNumberOfRows(); i++){
			ArrayList<DataBean> arSub1 = new ArrayList<DataBean>();
			DataBean dbeanObj = new DataBean();
			dbeanObj.setItemNumber(String.valueOf(originalWb.getSheet("NOV").getRow(i).getCell(0).getStringCellValue()));
			dbeanObj.setDescription(String.valueOf(originalWb.getSheet("NOV").getRow(i).getCell(1).getStringCellValue()));
			arSub1.add(dbeanObj);
			arMain.addAll(arSub1);
		}
		DataBean obj = null;
		
		for(int i=1;i<arMain.size();i++){
			obj = (DataBean) arMain.get(i);
			
			if(!obj.getItemNumber().equals(null)){
			Row r = wb.getSheet("NOV").createRow((short)i);
			Cell ce = r.createCell(0);
			ce.setCellValue(Integer.valueOf(obj.getItemNumber().trim()));
			Cell ce1 = r.createCell(1);
			ce1.setCellValue(obj.getDescription());
			}
		}

	}
	
	
	public void addDataToSheetDec(Workbook wb, Workbook originalWb){
		CellStyle style = wb.createCellStyle();
		style = wb.createCellStyle();
	    style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    style.setBorderBottom(CellStyle.BORDER_THICK);
	    style.setBorderTop(CellStyle.BORDER_THICK);
	    style.setBorderLeft(CellStyle.BORDER_THICK);
	    style.setBorderRight(CellStyle.BORDER_THICK);
		Row row = wb.getSheet("DEC").createRow((short)0);
		Cell cell = row.createCell(0);
		cell.setCellValue("Item (SKU)");
		wb.getSheet("DEC").autoSizeColumn(0);
		cell.setCellStyle(style);
		Cell cell1 = row.createCell(1);
		cell1.setCellValue("Description");
		wb.getSheet("DEC").autoSizeColumn(1);
		wb.getSheet("DEC").setColumnWidth(1, 50*256);
		cell1.setCellStyle(style);
		
		Calendar cal = Calendar.getInstance();
		cal.set(Calendar.YEAR, 2014);
	    cal.set(Calendar.MONTH, 11);
	    cal.set(Calendar.DAY_OF_MONTH, 1);
	    int maxDay = cal.getActualMaximum(Calendar.DAY_OF_MONTH);
	    SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
	    for (int i = 1; i <= maxDay; i++) {
	        Cell ceDate = row.createCell(i+1);
	    	ceDate.setCellValue(df.format(cal.getTime()));
			wb.getSheet("DEC").autoSizeColumn(i+1);
			ceDate.setCellStyle(style);
			cal.set(Calendar.DAY_OF_MONTH, i + 1);
	    }
		
		ArrayList<DataBean> arMain = new ArrayList<DataBean>();
		for(int i=2; i< originalWb.getSheet("DEC").getPhysicalNumberOfRows(); i++){
			ArrayList<DataBean> arSub1 = new ArrayList<DataBean>();
			DataBean dbeanObj = new DataBean();
			dbeanObj.setItemNumber(String.valueOf(originalWb.getSheet("DEC").getRow(i).getCell(0).getStringCellValue()));
			dbeanObj.setDescription(String.valueOf(originalWb.getSheet("DEC").getRow(i).getCell(1).getStringCellValue()));
			arSub1.add(dbeanObj);
			arMain.addAll(arSub1);
		}
		DataBean obj = null;
		
		for(int i=1;i<arMain.size();i++){
			obj = (DataBean) arMain.get(i);
			
			if(!obj.getItemNumber().equals(null)){
			Row r = wb.getSheet("DEC").createRow((short)i);
			Cell ce = r.createCell(0);
			ce.setCellValue(Integer.valueOf(obj.getItemNumber().trim()));
			Cell ce1 = r.createCell(1);
			ce1.setCellValue(obj.getDescription());
			}
		}

	}
	
}
