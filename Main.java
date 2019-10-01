package ExcelIFle.ExcelFIle;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.stream.Collectors;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xdgf.usermodel.XDGFCell;
import org.apache.poi.xssf.streaming.SXSSFRow.CellIterator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class lop {

	public static void main(String[] srgs) throws IOException {
		ArrayList<Integer> Abook = new ArrayList<Integer>();
		ArrayList<String> Bbook = new ArrayList<String>();
		ArrayList<String> Bbook1 = new ArrayList<String>();
		ArrayList<String> cbook = new ArrayList<String>();

		Map<String, String> mA = new HashMap<>();
		Map<String, String> mB = new HashMap<>();
		Map<String, String> mC = new HashMap<>();
		Map<String, String> md = new HashMap<>();
		Map<String, String> mE = new HashMap<>();
		Map<String, String> mf = new HashMap<>();

		File excelFile2 = new File("Path Of file");
		FileInputStream fis2 = new FileInputStream(excelFile2);
		XSSFWorkbook workbook2 = new XSSFWorkbook(fis2);
		XSSFSheet sheet2 = workbook2.getSheetAt(0);
		int rowStart = sheet2.getFirstRowNum();
		int rowEnd = sheet2.getLastRowNum();

		File excelFile = new File("Path Of file");
		FileInputStream fis = new FileInputStream(excelFile);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		int rowCount1 = sheet.getPhysicalNumberOfRows();

		File excelFile1 = new File("Path Of file");
		FileInputStream fis1 = new FileInputStream(excelFile1);
		XSSFWorkbook workbook1 = new XSSFWorkbook(fis1);
		XSSFSheet sheet1 = workbook1.getSheetAt(0);
		int rowCount2 = sheet1.getPhysicalNumberOfRows();

		for (int i = 0; i < rowCount1; i++) {
			XSSFRow row1 = sheet.getRow(i);

			String idstr1 = "";
			String idstr4 = "";
			String idstr5 = "";
			String idstr6 = "";
			XSSFCell id1 = row1.getCell(1);
			XSSFCell id4 = row1.getCell(12);
			XSSFCell id5 = row1.getCell(13);
			XSSFCell id6 = row1.getCell(14);

			if (id1 != null) {
				id1.setCellType(CellType.STRING);
				idstr1 = id1.getStringCellValue();

				idstr4 = id4.getStringCellValue();
				idstr5 = id5.getStringCellValue();
				idstr6 = id6.getStringCellValue();

				md.put(idstr1.toUpperCase(), idstr4);
				mA.put(idstr1.toUpperCase(), idstr5);
				mB.put(idstr1.toUpperCase(), idstr6);

			}
		}

		for (int j1 = 0; j1 < rowCount2; j1++) {
			XSSFRow row2 = sheet1.getRow(j1);

			String idstr2 = "";
			String idstr9 ="";
			String idstrr="";
			XSSFCell id1 = row2.getCell(0);
			XSSFCell id9 = row2.getCell(1);
			XSSFCell id7 = row2.getCell(5);

			if (id1 != null) {

				id1.setCellType(CellType.STRING);



				idstr2 = id1.getStringCellValue();
				idstr9 = id9.getStringCellValue();
				idstrr = id7.getStringCellValue();

              				
				mC.put(idstr2.toUpperCase(), idstr9);
				mE.put(idstr2.toUpperCase(), idstrr);
				Bbook.add(idstr2.toUpperCase());
			}
		}



		String opl = "";
		String op2 = "";

		int r = 1;

		for (Map.Entry m : mA.entrySet()) {

			opl = (String) m.getKey();
			op2 = (String) m.getValue();
			if (Bbook.contains(opl)) {
                sheet2.createRow(r);
				sheet2.getRow(r).createCell(1).setCellValue(opl);
				sheet2.getRow(r).createCell(5).setCellValue(op2);

				r++;
//				System.out.println("Keys: " + m.getKey() + ": Value:// " + m.getValue());

			}
			FileOutputStream forw = new FileOutputStream(excelFile2);
			workbook2.write(forw);

		}

		int p = 1;
		String op3 = "";
		String op4 = "";
		for (Map.Entry m : mB.entrySet()) {

			op4 = (String) m.getKey();
			op3 = (String) m.getValue();
			if (Bbook.contains(op4)) {

				sheet2.getRow(p).createCell(2).setCellValue(op3);
                 cbook.add(op4.toUpperCase());
				p++;
				System.out.println("Keys: " + m.getKey() + ": Value:// " + m.getValue());
 
			}
			FileOutputStream forw = new FileOutputStream(excelFile2);
			workbook2.write(forw);

		}
		System.out.println(cbook);

		int y = 1;
		String op7 = "";
		String op8 = "";
		for (Map.Entry m : md.entrySet()) {

			op7 = (String) m.getKey();
			op8 = (String) m.getValue();
			if (Bbook.contains(op7)) {

				sheet2.getRow(y).createCell(3).setCellValue(op8);

				y++;
				System.out.println("Keys: " + m.getKey() + ": Value:// " + m.getValue());

			}
			FileOutputStream forw = new FileOutputStream(excelFile2);
			workbook2.write(forw);

		}

	
		String op56="";
		String ophk= "hello";
		String op252="";
		int b=1;
		
for (int yo=0; yo<cbook.size();yo++){	
				String hu= cbook.get(yo);
				for (Map.Entry ma : mC.entrySet()){
					op56 = (String) ma.getKey();
					op252 = (String) ma.getValue();
				
		if (hu.equals(op56)) {
                
			op56=(String) ma.getValue();
				sheet2.getRow(b).createCell(7).setCellValue(op252);

				FileOutputStream forw = new FileOutputStream(excelFile2);
				workbook2.write(forw);
					b++;
				System.out.println("Keys: " + ma.getKey() + ": Value:// " + ma.getValue());
				
			}
		
				}
			}

String opq56="";
String ophqk= "hello";
String op2512="";
int l=1;

for (int o=0; o<cbook.size();o++){	
		String hu= cbook.get(o);
		for (Map.Entry ma : mE.entrySet()){
			opq56 = (String) ma.getKey();
			op2512 = (String) ma.getValue();
		
if (hu.equals(opq56)) {
        
	opq56=(String) ma.getValue();
		sheet2.getRow(l).createCell(6).setCellValue(op2512);

		FileOutputStream forw = new FileOutputStream(excelFile2);
		workbook2.write(forw);
			l++;
		System.out.println("Keys: " + ma.getKey() + ": Value:// " + ma.getValue());
		
	}

		}
	}
		
		

		System.out.println("SuccessFul");
	
		
		

	}
}
