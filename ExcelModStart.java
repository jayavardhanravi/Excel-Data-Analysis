
package maid.mycontent.orig;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import maid.myapi.com.CreateModules;
import maid.myapi.com.DataAnalysisModule;

public class ExcelModStart {
	
	public String cleanDataBase(){
		DataAnalysisModule daMod = new DataAnalysisModule();
		String a=null;
		if(daMod.cleanDataBaseIndaMod())
			a = "Success : Database has been cleaned";
		else
			a = "Failed : Unable to clean Database";
		return a;
	}
	
	public String ExcelMethodStart(String filePath,String dataFilePath){
		System.out.println("Received the File Names : ");
		System.out.println("Excel File Name : "+filePath);
		System.out.println("Excel Data File Name : "+dataFilePath);
		String a = null;
		 XSSFWorkbook workbook = null;
		 XSSFWorkbook originalWb = null;
		 XSSFWorkbook originalDataWb = null;
		CreateModules cmVar = new CreateModules();
		DataAnalysisModule daVar = new DataAnalysisModule();
		String bVar = cmVar.createBlankExcelSheet(filePath);
		File dataFileOriginal = new File(dataFilePath);
	      FileInputStream fIPDataOriginal;
		try {
			fIPDataOriginal = new FileInputStream(dataFileOriginal);
		    originalDataWb = new XSSFWorkbook(fIPDataOriginal);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			a = "ERROR Code JCat1001 : Error to open original data Excel file.";
			return a;
		}
	      if(dataFileOriginal.isFile() && dataFileOriginal.exists())
	      {
	         a = "Original Excel Data file oped successfully";
	      }
	      else
	      {
	         a = "Error to open original Excel Data file.";
	      }
		File fileOriginal = new File(filePath);
	      FileInputStream fIPOriginal;
		try {
			fIPOriginal = new FileInputStream(fileOriginal);
		    originalWb = new XSSFWorkbook(fIPOriginal);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			a = "ERROR Code JCat1001 : Error to open original Excel file.";
			return a;
		}
	      if(fileOriginal.isFile() && fileOriginal.exists())
	      {
	         a = "Original Excel file oped successfully";
	      }
	      else
	      {
	         a = "Error to open original Excel file.";
	      }
		File file = new File(bVar);
	      FileInputStream fIP;
		try {
			fIP = new FileInputStream(file);
		    workbook = new XSSFWorkbook(fIP);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			a = "ERROR Code JCat1002 : Error occured while opening the created file";
			return a;
		}
	      if(file.isFile() && file.exists())
	      {
	         a = "Excel file open successfully.";
	         cmVar.createSheetsInExcel(workbook);
	         daVar.ReadData(workbook,originalDataWb);
	         cmVar.addDataToSheetJan(workbook,originalWb);
	         cmVar.addDataToSheetFeb(workbook,originalWb);
	         cmVar.addDataToSheetMar(workbook,originalWb);
	         cmVar.addDataToSheetApr(workbook,originalWb);
	         cmVar.addDataToSheetMay(workbook,originalWb);
	         cmVar.addDataToSheetJun(workbook,originalWb);
	         cmVar.addDataToSheetJul(workbook,originalWb);
	         cmVar.addDataToSheetAug(workbook,originalWb);
	         cmVar.addDataToSheetSep(workbook,originalWb);
	         cmVar.addDataToSheetOct(workbook,originalWb);
	         cmVar.addDataToSheetNov(workbook,originalWb);
	         cmVar.addDataToSheetDec(workbook,originalWb);
	         System.out.println("Updated all the sheets with row details");
	      }
	      else
	      {
	         a = "Error to open Excel file.";
	      }
		if(bVar!=null)
			a = "Analysis Complete!!!!!  - File created successfully";
		else
			a = "File not created";
		FileOutputStream out = null;
		FileOutputStream out1 = null;
		FileOutputStream out2 = null;
		try {
			out = new FileOutputStream(bVar);
			workbook.write(out);
			out.close();
			workbook.close();
			daVar.analyseData(bVar);
			out1 = new FileOutputStream(filePath);
			originalWb.write(out1);
			out1.close();
			originalWb.close();
			out2 = new FileOutputStream(dataFilePath);
			originalDataWb.write(out2);
			out2.close();
			originalDataWb.close();
		}
		catch(Exception e)
		{
			e.printStackTrace();
			a = "ERROR Code JCat1003 : Error occured while closing the files";
			return a;
		}
		System.out.println("Process completed. Check the file..... [SUCCESS]");
		return a;
	}

}
