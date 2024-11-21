package com.scratch.GuviMavenProject;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadusingApachePOI {

	public static void main(String[] args) throws IOException {
		readexcel();

	}
    
	public static void readexcel() throws IOException {
		
		FileInputStream myfile = new FileInputStream("C:\\Users\\prabh\\Desktop\\data.xlsx");
		XSSFWorkbook myworkbook = new XSSFWorkbook(myfile);
		Sheet mysheet = myworkbook.getSheetAt(0);
		for (Row row :mysheet)
		{
			for (Cell cell : row) {
				switch (cell.getCellType()) {
				case STRING:
					System.out.print(cell.getRichStringCellValue()+"\t");
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue()+"\t");
					break;
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue()+"\t");
					break;
					default:
						System.out.println("UNKNOWN\t");
						break;
				}
				
			}
			System.out.println();
					}
		
			}
			
			
		
	}

