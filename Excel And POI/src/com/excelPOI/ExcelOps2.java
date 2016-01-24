package com.excelPOI;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;

public class ExcelOps2 {
	
	public static void main(String[] args){
		
		String[] items={"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","1","2","3","4"};
		Scanner scn=new Scanner(System.in);
		System.out.println("Enter the number of items to be in a row : ");
		int numPerRow=scn.nextInt();
				
		XSSFWorkbook wbook= new XSSFWorkbook();
		XSSFSheet sheet1=wbook.createSheet("MySheet1");
		XSSFRow row;
		XSSFCell cell;
		int numItems=items.length;
		int numRows=numItems/numPerRow;
		int counter=0;
		while (counter<numItems){
			for(int i=0;i<numRows;i++){
				row=sheet1.createRow(i);
				for(int j=0;j<numPerRow;j++){
					cell=row.createCell(j);
					cell.setCellValue(items[counter]);
					counter++;
				}
				
			}
		}
		try{
			File file1=new File("E:\\MyTestFiles\\TestExcel3.xlsx");	
			FileOutputStream fos=new FileOutputStream(file1);
		    wbook.write(fos);
		    System.out.println("Write Passed");
		}
		catch(IOException e){
			e.printStackTrace();
			System.out.println("Write Failed");
		}
	}

}
