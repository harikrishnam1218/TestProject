package com.practice;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.Iterator;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel2DatabaseTest {
	 
public static void main(String[] args) {
			//String name;
			Integer age;
			String name ,qualification, gender;
     try {
         FileInputStream file = new FileInputStream(new File("D:\\Harikrishna\\test.xlsx"));
         XSSFWorkbook workbook = new XSSFWorkbook(file);
         XSSFSheet sheet = workbook.getSheetAt(0);
         Iterator<Row> rowIterator = sheet.iterator();
         while(rowIterator.hasNext())
			{
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					cell.setCellType(CellType.STRING);
				}

				name = row.getCell(0).getStringCellValue();
				age = Integer.valueOf(row.getCell(1).getStringCellValue());
				qualification = row.getCell(2).getStringCellValue();
				gender = row.getCell(3).getStringCellValue();
				System.out.println("Reading Values from Excel  Name - " + name + ", Age - " + age + ", Qualification - "
						+ qualification+ ", Gender - " + gender);
				InsertRowInDB(name, age, qualification,gender);
			}
         file.close();
         
     } catch (FileNotFoundException e) {
         e.printStackTrace();
     } catch (IOException e) {
         e.printStackTrace();
     } catch (SQLException e) {
		e.printStackTrace();
	}
   }
     public static void InsertRowInDB(String name,Integer age,String qualification,String gender) throws SQLException{

    	 try {
             Properties properties = new Properties();
             properties.setProperty("user", "root");
             properties.setProperty("password", "Amma1991@");
             properties.setProperty("useSSL", "false");
             properties.setProperty("autoReconnect", "true");
             Class.forName("com.mysql.jdbc.Driver");
             Connection connect = DriverManager.getConnection("jdbc:mysql://localhost:3306/flightdata", properties);
             PreparedStatement ps = null;
             String sql = "Insert into Employee(name,age,qualification,gender) values(?,?,?,?)";
             ps = connect.prepareStatement(sql);
             ps=connect.prepareStatement(sql); 
             ps.setString(1, name);
             ps.setInt(2, age);
             ps.setString(3, qualification);
             ps.setString(4, gender);
             ps.executeUpdate();
             connect.close();
         } catch (ClassNotFoundException | SQLException e) {
             e.printStackTrace();
         }
		
     System.out.println("Values Inserted Successfully");
     }
}
