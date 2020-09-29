import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel3 {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		ArrayList<String> finalOutput = getData("mytest", "AddProfile");

		int size = finalOutput.size();

		for (int i = 0; i < size; i++) {

			System.out.println(finalOutput.get(i));
		}

	}

	public static ArrayList<String> getData(String sheetName, String testcaseName) {
		// fileInputStream argument
		ArrayList<String> data = new ArrayList<String>();

		try {

			FileInputStream fis = new FileInputStream("C://Users//aravindkoduri//Desktop//Test.xlsx");
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheet(sheetName);
			int lastRow = sheet.getLastRowNum();// Starts with index of 0
			int lastCell = sheet.getRow(0).getLastCellNum();// Gives the exact value

			for (int i = 1; i <= lastRow; i++) {

				XSSFRow row = sheet.getRow(i);

				for (int j = 0; j < lastCell; j++) {

					XSSFCell cell = row.getCell(j);

					data.add(cell.getStringCellValue());

					if (cell.getStringCellValue().equalsIgnoreCase("testcaseName")) {

						data.add(cell.getStringCellValue());
					}
				}
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return data;
	}

}
