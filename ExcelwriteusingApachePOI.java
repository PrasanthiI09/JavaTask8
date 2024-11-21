package com.scratch.GuviMavenProject;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

public class ExcelwriteusingApachePOI {

	public static void main(String[] args) {
		
		
		try (XSSFWorkbook workbook=new XSSFWorkbook()){

						//goto sheet 
			Sheet sheet=workbook.createSheet("Sheet1");  
			
			Object[][] data= {
					{"Name","Age","City"},
					{"pavan",45,"Bengaluru"},
					{"Priya",30,"Chennai"},
					{"Ajay",26,"Delhi"},
					{"Parmeshwari",26,"Pune"}
			};
			
		//write data to sheet
			int rowNum=0;
			for(Object[]  rowdata:data) {
				Row row=sheet.createRow(rowNum++);
				int colNum=0;//inner column is used for printing the data
				for(Object field:rowdata) {
					Cell cell=row.createCell(colNum++);
					if(field instanceof String) {
						cell.setCellValue((String)field);
					}else if(field instanceof Integer) {
						cell.setCellValue((Integer)field);
					}
				}
			}
			
			//workbook to an excel
			try(FileOutputStream os=new FileOutputStream("dataa.xlsx")){
				workbook.write(os);
			}
			
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
