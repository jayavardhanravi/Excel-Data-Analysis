
package maid.myapi.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Calendar;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataAnalysisModule {
	private PreparedStatement preparedStatement = null;
	private ResultSet rs = null;
	
	public boolean cleanDataBaseIndaMod(){
		Connection con1 = DbConnection.getConnection();
		Connection con2 = DbConnection.getConnection();
		String dropTableQuery = "drop table DataSheet_tbl;";
		String createTableQuery = "create table DataSheet_tbl(item_id int not null, item_date date , item_quantity int, primary key (item_id,item_date));";
		try {
			preparedStatement = con1.prepareStatement(dropTableQuery);
			preparedStatement.execute();
			DbConnection.closePreparedStatement(preparedStatement);
			DbConnection.closeConnection(con1);
			preparedStatement = con2.prepareStatement(createTableQuery);
			preparedStatement.execute();
			DbConnection.closePreparedStatement(preparedStatement);
			DbConnection.closeConnection(con2);
			return true;
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return false;
		}
	}
	
	public void ReadData(Workbook wb, Workbook dataWb){
		ReadDataToArray(dataWb);
		
	}
	
	private void ReadDataToArray(Workbook dataWb){
		Connection con = DbConnection.getConnection();
		String DataSheetInsertQuery = "insert into DataSheet_tbl values(?,?,?);";
		for(int i=1;i<dataWb.getSheet("Data").getPhysicalNumberOfRows();i++){
			if(!dataWb.getSheet("Data").getRow(i).getCell(0).getDateCellValue().equals(null)){
				try {
					preparedStatement = con.prepareStatement(DataSheetInsertQuery);
					preparedStatement.setString(1, String.valueOf(dataWb.getSheet("Data").getRow(i).getCell(1).getNumericCellValue()));
					preparedStatement.setDate(2, new Date((dataWb.getSheet("Data").getRow(i).getCell(0).getDateCellValue()).getTime()));
					preparedStatement.setString(3, String.valueOf(dataWb.getSheet("Data").getRow(i).getCell(2).getNumericCellValue()));
					preparedStatement.executeUpdate();
				} catch (SQLException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			
			}
			
		}
		DbConnection.closePreparedStatement(preparedStatement);
		DbConnection.closeConnection(con);	
		System.out.println("All the data in the sheet is analysed and updated to memory");
	}
	
	public void analyseData(String createdFileName){
		Connection con = DbConnection.getConnection();
		String DataSheetSelectQuery = "select item_quantity from DataSheet_tbl where item_id = ? and item_date = ?;"; 
		XSSFWorkbook createdWorkbook = null;
		File dataFileOriginal = new File(createdFileName);
	      FileInputStream fIPCreatedFile;
		try {
			fIPCreatedFile = new FileInputStream(dataFileOriginal);
			createdWorkbook = new XSSFWorkbook(fIPCreatedFile);
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		for(int k=0;k<12;k++){
			for(int i=1;i<createdWorkbook.getSheetAt(k).getPhysicalNumberOfRows();i++){
				double itemNumber = createdWorkbook.getSheetAt(k).getRow(i).getCell(0).getNumericCellValue();
				Calendar cal = Calendar.getInstance();
				cal.set(Calendar.YEAR, 2014);
			    cal.set(Calendar.MONTH, k);
			    int maxDay = cal.getActualMaximum(Calendar.DAY_OF_MONTH);
			    for(int j=1;j<=maxDay;j++){
			    	java.util.Date itemDate = Date.valueOf(createdWorkbook.getSheetAt(k).getRow(0).getCell(j+1).getStringCellValue());
			    	java.sql.Date itemDateSql = new java.sql.Date(itemDate.getTime());
			    	try {
						preparedStatement = con.prepareStatement(DataSheetSelectQuery);
						preparedStatement.setDouble(1, itemNumber);
						preparedStatement.setDate(2, itemDateSql);
						rs = preparedStatement.executeQuery();
						while(rs.next()){
							double quant = rs.getDouble("item_quantity");
							createdWorkbook.getSheetAt(k).getRow(i).createCell(j+1).setCellValue(quant);
						}
					} catch (SQLException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
			    }
				
			}
		}
		DbConnection.closePreparedStatement(preparedStatement);
		DbConnection.closeConnection(con);
		FileOutputStream out;
		try {
			out = new FileOutputStream(createdFileName);
			createdWorkbook.write(out);
			out.close();
			System.out.println("Data analysed and created the excel sheet.");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
