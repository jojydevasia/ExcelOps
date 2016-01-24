package com.excelPOI;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.ss.usermodel.Cell;

public class ExcelOps3 {
	
	public static void main(String[] args){
		String [] cellContents={"Apple","Ball","Cat","Dog","Egg","Fish","Grapes","Horse","Ink","Jackal","King",
				             "Lion","Monkey","Nose","Orange","Potato","Queen","Rainbow","Star","Tortoise",
				             "Umbrella","Vulture","Wolf","Xenon","Yatch","Zebra", "Apricot","Brocoli",
				             "Carrot"};
		System.out.println("ArrayLength: "+cellContents.length);
		
		
		XSSFWorkbook wbook1=new XSSFWorkbook();
		XSSFSheet sheet1=wbook1.createSheet();
		XSSFSheet sheet2=wbook1.createSheet("MyNewSheet");
		XSSFRow row=null;
		Cell cell=null;
		for(String k:cellContents){
			System.out.print(k+" , ");
		}
		int x=cellContents.length;
		int counter=0;
		for (int j=0;j<6;j++){
		row=sheet1.createRow(j);
		for(int m=0;m<5;m++){
			
		cell=row.createCell(m);
		cell.setCellValue(cellContents[counter]);
		counter++;
		}
		}
		
		try{
			File file1=new File("E:\\MyTestFiles\\NewTestExcel.xlsx");
			FileOutputStream fos= new FileOutputStream(file1);
					
			FileInputStream fis=new FileInputStream(file1);
			
			//wbook1.cloneSheet(1);
			//wbook1.createSheet("NewSheet1");
			
			//wbook1.setSheetName(0, "Latest Sheet");
			
			//wbook1.removeSheetAt(0);
			System.out.println("Number of sheets: "+wbook1.getNumberOfSheets());
			System.out.println("SheetName: "+wbook1.getSheetName(0));
			
			wbook1.write(fos);
			
		}
		catch(IOException e){
			e.printStackTrace();
		}
		
		
	}

}

