package com.excelPOI;

import java.util.Scanner;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;

public class CompareSpreadSheet {
	public static void main(String[] args){
		try{
			File file1=new File("E:\\MyTestFiles\\TestExcel1.xlsx");
			File file2=new File("E:\\MyTestFiles\\TestExcel2.xlsx");
			
			FileInputStream exlFis1=new FileInputStream(file1);
			FileInputStream exlFis2=new FileInputStream(file2);
			
			XSSFWorkbook workbook1=new XSSFWorkbook(exlFis1);
			XSSFWorkbook workbook2=new XSSFWorkbook(exlFis2);
			
			XSSFSheet sheet1=workbook1.getSheetAt(0);
			XSSFSheet sheet2=workbook2.getSheetAt(0);
			
			if(compareTwoSheets(sheet1,sheet2)){
				System.out.println("The two sheets are equal");
			}else{
				System.err.println("Sheets are not equal");
			}
			exlFis1.close();
			exlFis2.close();
			
						
		}catch(Exception e){
			e.printStackTrace();
			System.out.println("Operation Failed. Check The Code Buddy :-D :-D :-D");
		}
		
	}
	
	public static boolean compareTwoSheets(XSSFSheet sheet1, XSSFSheet sheet2){
		boolean equalSheets=true;
		//int firstRowNum1=sheet1.getFirstRowNum();
		//int lastRowNum1=sheet1.getLastRowNum();
		//int numRows1=(lastRowNum1-firstRowNum1)+1;
		//System.out.println("NUmRows: "+numRows1);
		int numRows1=sheet1.getPhysicalNumberOfRows();
		int numRows2=sheet2.getPhysicalNumberOfRows();
		if(numRows1==numRows2){
			for(int i=0;i<numRows1;i++){
				XSSFRow row1=sheet1.getRow(i);
				XSSFRow row2=sheet2.getRow(i);
				
				if(!compareTwoRows(row1,row2)){
					equalSheets=false;
					System.out.println("Row "+ i+" Not Equal");
					break;
				}else {
					System.out.println("Row "+i+" Is Equal");
				}
			}
		}else {
			equalSheets=false;
			System.out.println("Diff Num of Rows");
		
		}
		
		return equalSheets;
	}
	
	
	public static boolean compareTwoRows(XSSFRow row1, XSSFRow row2){
		boolean equalRows=true;
		int numCells1=row1.getPhysicalNumberOfCells();
		int numCells2=row2.getPhysicalNumberOfCells();
		if(numCells1==numCells2){
			for(int j=0;j<numCells1;j++){
				XSSFCell cell1=row1.getCell(j);
				XSSFCell cell2=row2.getCell(j);
				
				if(!compareTwoCells(cell1,cell2)){
					equalRows=false;
					System.out.println("Cell "+j+" Not Equal");
					break;
				}else{
					System.out.println("Cell "+j+" Is Equal");
				}
				
			}
		}else equalRows=false;
		
		return equalRows;
	}
	
	
	public static boolean compareTwoCells(XSSFCell cell1, XSSFCell cell2){
		boolean equalCells=true;
		
		
		return equalCells;
	}

}
