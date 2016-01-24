package com.excelPOI;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.FileNotFoundException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;


public class RWExcel {
	
	public static void main(String[] args){
		XSSFWorkbook workbook= new XSSFWorkbook();
		XSSFSheet spreadsheet=workbook.createSheet("MySheet1");
		
		for(int i=0;i<10;i++){
			XSSFRow row=spreadsheet.createRow(i);
			for(int j=0;j<5;j++){
				Cell cell=row.createCell(j);
				cell.setCellValue("Cell " + i+" - "+j);
			}
		}
		
		try {
			FileOutputStream fout=new FileOutputStream(new File("C:\\Users\\jojydevasia\\Desktop\\MyNewWorkBook1.xslx"));
			workbook.write(fout);
			workbook.close();
			fout.close();
			
			
			FileInputStream fin=new FileInputStream(new File("C:\\Users\\jojydevasia\\Desktop\\MyNewWorkBook1.xslx"));
			XSSFWorkbook  wb= new XSSFWorkbook(fin);
			XSSFSheet sht=wb.getSheet("MySheet1");
			Iterator<Row> itr= sht.iterator();
			
			while(itr.hasNext()){
				Row nextRow=itr.next();
				Iterator<Cell> cellItr=nextRow.cellIterator();
				
				while(cellItr.hasNext()){
					Cell cel=cellItr.next();
					
					switch(cel.getCellType()){
					
					case Cell.CELL_TYPE_STRING:
						System.out.println(cel.getStringCellValue());
						break;
					case Cell.CELL_TYPE_NUMERIC:
						System.out.println(cel.getNumericCellValue());
						break;
					case Cell.CELL_TYPE_BLANK:
						System.out.println(cel.getStringCellValue());
						break;
					case Cell.CELL_TYPE_BOOLEAN:
						System.out.println(cel.getBooleanCellValue());
						break;
					}
					System.out.print(" - ");
				}
				System.out.println();
			}
			wb.close();
			fin.close();
			
			
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

}
