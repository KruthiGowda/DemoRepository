package com.sgtesting.excelfileassignments;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class VegetableNames 
{
	public static void main(String[] args) 
	{
		createfolder();
		writecontent();
	}
	private static void createfolder()
	{
		try
		{
			File f = new File("F:\\Vegetables");
			boolean a = f.mkdir();
			System.out.println(a);
			File f1 = new File("F:\\Vegetables\\Vegetables.xlsx");
			boolean a1 = f1.createNewFile();
			System.out.println(a1);
		}catch(Exception e)
		{
			e.printStackTrace();
		}
	}
	private static void writecontent()
	{
		Workbook wb = null;
		Sheet sh = null;
		Row row = null;
		Cell cell = null;
		FileOutputStream fout = null;
		try
		{
			wb = new XSSFWorkbook();
			sh = wb.createSheet("Sheet 1");
			row = sh.createRow(0);
			cell = row.createCell(0);
			cell.setCellValue("Vegetable 1");
			
			row = sh.createRow(1);
			cell = row.createCell(1);
			cell.setCellValue("Vegetable 2");
			
			row = sh.createRow(2);
			cell = row.createCell(2);
			cell.setCellValue("Vegetable 3");
			
			row = sh.createRow(3);
			cell = row.createCell(3);
			cell.setCellValue("Vegetable 4");
			
			row = sh.createRow(4);
			cell = row.createCell(4);
			cell.setCellValue("Vegetable 5");
			
			row = sh.createRow(5);
			cell = row.createCell(5);
			cell.setCellValue("Vegetable 6");
			
			row = sh.createRow(6);
			cell = row.createCell(6);
			cell.setCellValue("Vegetable 7");
			
			row = sh.createRow(7);
			cell = row.createCell(7);
			cell.setCellValue("Vegetable 8");
			
			row = sh.createRow(8);
			cell = row.createCell(8);
			cell.setCellValue("Vegetable 9");
			
			row = sh.createRow(9);
			cell = row.createCell(9);
			cell.setCellValue("Vegetable 10");
			
			row = sh.createRow(10);
			cell = row.createCell(10);
			cell.setCellValue("Vegetable 11");
			
			row = sh.createRow(11);
			cell = row.createCell(11);
			cell.setCellValue("Vegetable 12");
			
			row = sh.createRow(12);
			cell = row.createCell(12);
			cell.setCellValue("Vegetable 13");
			
			row = sh.createRow(13);
			cell = row.createCell(13);
			cell.setCellValue("Vegetable 14");
			
			row = sh.createRow(14);
			cell = row.createCell(14);
			cell.setCellValue("Vegetable 15");
			
			row = sh.createRow(15);
			cell = row.createCell(15);
			cell.setCellValue("Vegetable 16");
			
			row = sh.createRow(16);
			cell = row.createCell(16);
			cell.setCellValue("Vegetable 17");
			
			row = sh.createRow(17);
			cell = row.createCell(17);
			cell.setCellValue("Vegetable 18");
			
			row = sh.createRow(18);
			cell = row.createCell(18);
			cell.setCellValue("Vegetable 19");
			
			row = sh.createRow(19);
			cell = row.createCell(19);
			cell.setCellValue("Vegetable 20");
			
			fout = new FileOutputStream("F:\\Vegetables\\Vegetables.xlsx");
			wb.write(fout);
		}catch(Exception e)
		{
			e.printStackTrace();
		}
		finally
		{
			try
			{
				fout.close();
				wb.close();
			}catch(Exception e)
			{
				e.printStackTrace();
			}
		}
	}
}

