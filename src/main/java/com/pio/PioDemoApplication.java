package com.pio;

import java.io.File;
import java.io.FileInputStream;

import java.io.FileOutputStream;

import java.text.SimpleDateFormat;

import java.util.Date;

import java.util.Iterator;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class PioDemoApplication {

	public static void main(String[] args) throws Exception 
	{
		SpringApplication.run(PioDemoApplication.class, args);
		
	    String path = System.getProperty("user.home") + "/Downloads/xeFile.xlsx";
	    FileInputStream file = new FileInputStream(new File(path));
	    
	    Workbook workbook = new XSSFWorkbook(file);
	    
	    Sheet sheet = workbook.getSheetAt(0);

	    Iterator<Row> rows = sheet.iterator();
	    int rowNumber = 0;
		double tot = 0;

	    while(rows.hasNext()) 
	    {
	    	Row currentRow = rows.next();
	    	Iterator<Cell> cellsInRow = currentRow.iterator();
	    	int cellIdx = 0;
	    	
	    	while(cellsInRow.hasNext()) 
	    	{
	    		Cell currentCell = cellsInRow.next();
	    		switch (cellIdx) 
	    		{
					case 0:
						if(DateUtil.isCellDateFormatted(currentCell)) {
							
							//System.out.println("Date 1"+currentCell.getDateCellValue());
							Date string = currentCell.getDateCellValue();
							
							final String OLD_FORMAT = "dd/MM/yyyy";
							final String NEW_FORMAT = "yyyy/dd/MM";

							// August 12, 2010
							Date oldDateString = string;
							String newDateString;

							SimpleDateFormat sdf = new SimpleDateFormat(OLD_FORMAT);
						
							sdf.applyPattern(NEW_FORMAT);
							newDateString = sdf.format(oldDateString);
							
							
							System.out.println(newDateString);//Mon Jul 05 00:00:00 IST 2010
							

						}
							
						else
						System.out.println("case 0 "+currentCell.getStringCellValue());
						break;
					case 1:
							System.out.println("Numeric "+currentCell.getNumericCellValue());
							tot += currentCell.getNumericCellValue();
						break;
					case 2:
						System.out.println("case 2 "+currentCell.getStringCellValue());
						break;
	
					default:
						break;
					}
	    		cellIdx++;
	    	}
	    }
	    
	    workbook.close();
	    System.out.println(tot);
	    
	    
	    
	    
	    Workbook workbook2 = new XSSFWorkbook();
	    Sheet sheet2 = workbook2.createSheet("personns");
	    
	    sheet2.setColumnWidth(0, 6000);
	    sheet2.setColumnWidth(1, 4000);
	    
	    Row header  = sheet2.createRow(0);
	    
	    CellStyle headerStyle = workbook2.createCellStyle();

	    
	    XSSFFont font = ((XSSFWorkbook) workbook2).createFont();
	    font.setFontName("Arial");
	    font.setFontHeightInPoints((short) 13);
	    font.setBold(true);
	    headerStyle.setFont(font);
	    
	    Cell headerCell = header.createCell(0);
	    headerCell.setCellValue("Name");
	    headerCell.setCellStyle(headerStyle);
	    
	    headerCell = header.createCell(1);
	    headerCell.setCellValue("Age");
	    headerCell.setCellStyle(headerStyle);
	    
	    CellStyle style = workbook2.createCellStyle();
	    style.setWrapText(true);
	    
	    Row row = sheet2.createRow(1);
	    Cell cell = row.createCell(0);
	    cell.setCellValue("John Smith");
	    cell.setCellStyle(style);
	    
	    cell = row.createCell(1);
	    cell.setCellValue(20);
	    cell.setCellStyle(style);
	    
	    row = sheet2.createRow(2);
	    cell = row.createCell(0);
	    cell.setCellValue("John Smith");
	    cell.setCellStyle(style);
	    
	    
	    
	    
	    String path2 = System.getProperty("user.home") + "/Downloads/temp.xlsx";
	    FileOutputStream outputStream = new FileOutputStream(path2);
	    workbook2.write(outputStream);
	    workbook2.close();
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	   
		/*
		 * int rows = 0, cols = 0; for (Row row : sheet) {
		 * System.out.println(row.getRowNum()); data.put(rows, new ArrayList<String>());
		 * for (Cell cell : row) { Cell cell2 =
		 * sheet.getRow(row.getRowNum()).getCell(1));
		 * 
		 * switch (cell.getCellType()) { case STRING :
		 * System.out.println(cell.getStringCellValue()); break; case NUMERIC :
		 * System.out.println(cell.getNumericCellValue()); break; case BOOLEAN :
		 * System.out.println(cell.getBooleanCellValue()); break;
		 * 
		 * } cols++; } System.out.println(); rows++; } System.out.println("Rows "+rows);
		 * System.out.println("Cols "+cols);
		 */
	    
	    
	   
	}

}
