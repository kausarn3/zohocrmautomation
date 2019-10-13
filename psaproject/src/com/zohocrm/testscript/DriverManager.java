package com.zohocrm.testscript;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;

public class DriverManager extends Keywords {

	static ArrayList data;
	static DriverManager dm = new DriverManager();

	static {
		try {
			data = new ArrayList();
			FileInputStream file = new FileInputStream(
					"F:\\WorkSpaces\\selenium\\zohocrm\\src\\com\\zohocrm\\testdata\\models.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					if (cell.getCellTypeEnum() == CellType.STRING) {
						data.add(cell.getStringCellValue());
					}
					if (cell.getCellTypeEnum() == CellType.NUMERIC) {
						data.add(cell.getNumericCellValue());
					}
					if (cell.getCellTypeEnum() == CellType.BOOLEAN) {
						data.add(cell.getBooleanCellValue());
					}

				}
			}

		} catch (Exception e) {

		}

	}

	@Test
	public void main() {
		for (int i = 0; i < data.size(); i++) {

			if (data.get(i).equals("Leads") && data.get(i + 1).equals("yes")) {
				try {
					ArrayList data1 = new ArrayList();
					FileInputStream file = new FileInputStream(
							"F:\\WorkSpaces\\selenium\\zohocrm\\src\\com\\zohocrm\\testdata\\Leads.xlsx");
					XSSFWorkbook workbook = new XSSFWorkbook(file);
					XSSFSheet sheet = workbook.getSheetAt(0);
					Iterator<Row> rowIterator = sheet.iterator();
					while (rowIterator.hasNext()) {
						Row row = rowIterator.next();
						Iterator<Cell> cellIterator = row.cellIterator();
						while (cellIterator.hasNext()) {
							Cell cell = cellIterator.next();
							if (cell.getCellTypeEnum() == CellType.STRING) {
								data1.add(cell.getStringCellValue());
							}
							if (cell.getCellTypeEnum() == CellType.NUMERIC) {
								data1.add(cell.getNumericCellValue());
							}
							if (cell.getCellTypeEnum() == CellType.BOOLEAN) {
								data1.add(cell.getBooleanCellValue());
							}
						}

					}

					for (int j = 0; j < data1.size(); j++) {
						if (data1.get(j).equals("Openbrowser")) {
							String Keywordname = (String) data1.get(j);
							String testData = (String) data1.get(j + 1);
							String objectname = (String) data1.get(j + 2);
							String runMode = (String) data1.get(j + 3);
							System.out.println(Keywordname + " " + testData + " " + objectname + " " + runMode);
							if (runMode.equals("yes")) {

								dm.openbrowser();
							}

						}
						if (data1.get(j).equals("Navigate")) {
							String Keywordname = (String) data1.get(j);
							String testData = (String) data1.get(j + 1);
							String objectname = (String) data1.get(j + 2);
							String runMode = (String) data1.get(j + 3);
							System.out.println(Keywordname + " " + testData + " " + objectname + " " + runMode);
							if (runMode.equals("yes")) {

								dm.navigate(testData);

							}
						}
						if (data1.get(j).equals("input")) {
							String Keywordname = (String) data1.get(j);
							String testData = (String) data1.get(j + 1);
							String objectname = (String) data1.get(j + 2);
							String runMode = (String) data1.get(j + 3);
							System.out.println(Keywordname + " " + testData + " " + objectname + " " + runMode);
							if (runMode.equals("yes")) {

								dm.input(testData, objectname);

							}
						}
						if (data1.get(j).equals("click")) {
							String Keywordname = (String) data1.get(j);
							String testData = (String) data1.get(j + 1);
							String objectname = (String) data1.get(j + 2);
							String runMode = (String) data1.get(j + 3);
							System.out.println(Keywordname + " " + testData + " " + objectname + " " + runMode);
							if (runMode.equals("yes")) {

								dm.click(objectname);

							}
						}

						if (data1.get(j).equals("type")) {
							String Keywordname = (String) data1.get(j);
							String testdata = (String) data1.get(j + 1);
							String objectname = (String) data1.get(j + 2);
							String runMode = (String) data1.get(j + 3);
							System.out.println(Keywordname + " " + testdata + " " + objectname + " " + runMode);
							if (runMode.equals("yes")) {
								dm.type(testdata);

							}
						}
						if (data1.get(j).equals("close")) {
							String Keywordname = (String) data1.get(j);
							String testdata = (String) data1.get(j + 1);
							String objectname = (String) data1.get(j + 2);
							String runMode = (String) data1.get(j + 3);
							System.out.println(Keywordname + " " + testdata + " " + objectname + " " + runMode);
							if (runMode.equals("yes")) {
								dm.close();

							}
						}

						if (data1.get(j).equals("gettitle")) {
							String Keywordname = (String) data1.get(j);
							String testdata = (String) data1.get(j + 1);
							String objectname = (String) data1.get(j + 2);
							String runMode = (String) data1.get(j + 3);
							System.out.println(Keywordname + " " + testdata + " " + objectname + " " + runMode);
							if (runMode.equals("yes")) {
								String actualdata = dm.gettitle();
								Assert.assertEquals(testdata, actualdata);

							}
						}
						if (data1.get(j).equals("verifybox")) {
							String Keywordname = (String) data1.get(j);
							String testdata = (String) data1.get(j + 1);
							String objectname = (String) data1.get(j + 2);
							String runMode = (String) data1.get(j + 3);
							System.out.println(Keywordname + " " + testdata + " " + objectname + " " + runMode);
							if (runMode.equals("yes")) {
								String actualdata = dm.verifybox(objectname);
								Assert.assertEquals(testdata, actualdata);

							}
						}

						if (data1.get(j).equals("waitStatement")) {
							String Keywordname = (String) data1.get(j);
							double testData = (double) data1.get(j + 1);
							String objectname = (String) data1.get(j + 2);
							String runMode = (String) data1.get(j + 3);
							System.out.println(Keywordname + " " + testData + " " + objectname + " " + runMode);
							if (runMode.equals("yes")) {

								dm.waitStatement(testData);

							}
						}

					}

				} catch (Exception e) {

				}
			}
		}
	}
}
