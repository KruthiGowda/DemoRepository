package com.sgtesting.excelfileassignments;

import java.io.File;
import java.io.FileOutputStream;
	import org.apache.poi.ss.usermodel.Cell;
	import org.apache.poi.ss.usermodel.Row;
	import org.apache.poi.ss.usermodel.Sheet;
	import org.apache.poi.ss.usermodel.Workbook;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	public class FruitNames 
	{
		public static void main(String[] args) 
		{
			createfolder();
			writeContent();
		}
		private static void createfolder()
		{
			try
			{
				File f = new File("F:\\Fruits");
				boolean a = f.mkdir();
				System.out.println(a);
				File f1 = new File("F:\\Fruits\\Fruits.xlsx");
				boolean a1 = f1.createNewFile();
				System.out.println(a1);
			}catch(Exception e)
			{
				e.printStackTrace();
			}
			
		}
		
		private static void writeContent()
		{
			Workbook wb=null;
			Sheet sh=null;
			Row row=null;
			Cell cell=null;
			FileOutputStream fout=null;
			try
			{
				wb=new XSSFWorkbook();
				sh=wb.createSheet("Sheet 1");
				row=sh.createRow(0);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 1");
				
				row=sh.createRow(1);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 2");
				
				row=sh.createRow(2);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 3");
				
				row=sh.createRow(3);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 4");
				
				row=sh.createRow(4);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 5");
				
				row=sh.createRow(5);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 6");
				
				row=sh.createRow(6);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 7");
				
				row=sh.createRow(7);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 8");
				
				row=sh.createRow(8);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 9");
				
				row=sh.createRow(9);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 10");
				
				row=sh.createRow(10);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 11");
				
				row=sh.createRow(11);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 12");
				
				row=sh.createRow(12);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 13");
				
				row=sh.createRow(13);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 14");
				
				row=sh.createRow(14);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 15");
				
				row=sh.createRow(15);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 16");
				
				row=sh.createRow(16);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 17");
				
				row=sh.createRow(17);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 18");
				
				row=sh.createRow(18);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 19");
				
				row=sh.createRow(19);
				cell=row.createCell(0);
				cell.setCellValue("Fruit 20");
				
				fout=new FileOutputStream(":\\Fruits\\Fruits.xlsx");
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
