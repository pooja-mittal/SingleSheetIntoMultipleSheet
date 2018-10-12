package com.divide.multisheet.SeprateSheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.collect.BiMap;
import com.google.common.collect.HashBiMap;

public class SeparateSheet {

	public static void main(String[] args) {
		readXLS();

	}

	public static void readXLS() {
		String sheetPath = "C:\\Users\\Pooja\\Desktop\\Fruits_sheet.xlsx";
		int a = 0, b = 0;
		Map<Integer, String> map = new LinkedHashMap<Integer, String>();
		BiMap<Integer, String> bimap = HashBiMap.create();
		try {
			File file = new File(sheetPath);
			FileInputStream fis = new FileInputStream(file);
			XSSFWorkbook excelWBook = new XSSFWorkbook(fis);
			XSSFSheet excelSheet = excelWBook.getSheetAt(0);
			int totalRows = excelSheet.getLastRowNum() + 1;
			System.out.println("total rows " + totalRows);
			
			Set<String> unique=new LinkedHashSet<String>();
			Set<String> duplicate=new LinkedHashSet<String>();
			
			Iterator<Row> rowIterator = excelSheet.iterator();
			String compareString=null;
			String cellData = null;
			for(int i=0; i<totalRows; i++) {
				Row row=excelSheet.getRow(i);
				for(int j=0; j<row.getLastCellNum(); j++) {
					Cell cell=row.getCell(j);
					DataFormatter formatter = new DataFormatter();
		            cellData= formatter.formatCellValue(cell);
		            compareString=cellData;
//		            System.out.print("ist element " + row.getCell(0).getStringCellValue()+ " row  "+ cellData+ " ");
		            System.out.print(i + " row " + j + " element "+ cellData);
		            System.out.println();
				}
				if((cellData.contains(row.getCell(0).getStringCellValue()))) {
					unique.add(cellData);
					System.out.println(" contain");
				}
				
//				
//				System.out.println(".set..");
//			    for(String s: unique) {
//			    	System.out.print(s + " ");
//			    	 System.out.println();
//			    	
//			    }
			}
				
//		    while(rowIterator.hasNext()) {
//		        Row row = rowIterator.next();		
//		        
//		        Iterator<Cell> cellIterator = row.cellIterator();
//		        
//		        
//		        while(cellIterator.hasNext()) {
//		            Cell cell = cellIterator.next();
//		            DataFormatter formatter = new DataFormatter();
//		            cellData= formatter.formatCellValue(cell);	
//		            System.out.println(cellData+ " size " + cellData.length());
////		            unique.add(cellData);
//		            
//		        }
//		        System.out.println("..........");
//					
//	        }
//		    System.out.println(".set..");
//		    for(String s: unique) {
//		    	System.out.print(s + " ");
//		    	 System.out.println();
//		    	
//		    }
		   

			for (int i = 0; i < totalRows; i++) {
				XSSFCell cell = excelSheet.getRow(i).getCell(0);
				DataFormatter formatter = new DataFormatter();
				String cellData1 = formatter.formatCellValue(cell);
				
				try {
					bimap.put(b, cellData1);
					b++;

				} catch (Exception e) {
					map.put(a, cellData1);
					a++;
				}
				// System.out.println("coloumn data "+cellData);
			}
			System.out.println("total unique fruits are " + bimap.size());
			for (BiMap.Entry<Integer, String> m : bimap.entrySet()) {
//				writeInMultipleSheet(m.getValue(),set);
				System.out.print("key " + m.getKey() + " value " + m.getValue());
				System.out.println();
			}
			

			// System.out.println("...............");
			// System.out.println("map size " + map.size());
			// for (Map.Entry<Integer, String> m : map.entrySet()) {
			// System.out.print("key " + m.getKey() + " value " + m.getValue());
			// System.out.println();
			// }

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static void writeInMultipleSheet(String sheetName, Set<String> cellData) {
		File file = null;
		OutputStream fos = null;
		XSSFWorkbook workbook = null;
		try {
			String sheetPath = "C:\\Users\\Pooja\\Desktop\\Fruits_sheet.xlsx";
			file = new File(sheetPath);
			Sheet sheet = null;
			if (file.exists()) {
				workbook = (XSSFWorkbook) WorkbookFactory.create(new FileInputStream(file));
				
			} else {
				workbook = new XSSFWorkbook();
			}
			sheet = workbook.createSheet(sheetName);
			Row currentRow = sheet.createRow(0);
			currentRow.createCell(0).setCellValue("Name");
			currentRow.createCell(1).setCellValue("Order ID");
			currentRow.createCell(2).setCellValue("To Email");
			currentRow.createCell(3).setCellValue("Code");
			currentRow.createCell(4).setCellValue("Pin");
			currentRow.createCell(5).setCellValue("Amount");
			currentRow.createCell(6).setCellValue("Validity Date");
			currentRow.createCell(7).setCellValue("Remaining Amount");
			currentRow.createCell(8).setCellValue("Current Validity");
			
			fos = new FileOutputStream(file);
            workbook.write(fos);
            fos.flush();
            System.out.println("multiple sheets are created");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	

}
