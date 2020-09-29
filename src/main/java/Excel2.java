import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.cigniti.utilities.Sheet;

public class Excel2 {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		ArrayList<String> finalOutput = getData("AddProfile");

		int size = finalOutput.size();

		for (int i = 0; i < size; i++) {

			System.out.println(finalOutput.get(i));
		}

	}

	public static void getData(String sheetName, String testcaseName) throws IOException {
		// fileInputStream argument
		ArrayList<String> a = new ArrayList<String>();

		FileInputStream fis = new FileInputStream("C://Users//aravindkoduri//Desktop//Test.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet(sheetName);
		XSSFRow row;
		int rows = sheet.getLastRowNum();
		int cols = 3;
		
		String[] tdata = null;
		int sheetsCount = workbook.getNumberOfSheets();
		
		for ( int i = 1; i < rows ; i++) {
			
			Iterator<Row> sheetRows = sheet.iterator();
			Row respectiveRowValue = sheetRows.next();
			
			if ( row.getCell(i).getStringCellValue().equalsIgnoreCase(""))
		}

		for (int i = 1; i < rows; i++) {
			if (sheet.getCell(0, i).getContents().equalsIgnoreCase("Yes")
					&& sheet.getCell(1, i).getContents().equalsIgnoreCase(testcaseName)) { // matching key;
				tdata = new String[cols];

				for (int j = 0; j < cols; j++) {

					if (sheet.getCell(j, i).getContents().equalsIgnoreCase("n/a")) {
						tdata[j] = "";
					} else {
						tdata[j] = sheet.getCell(j, i).getContents();
					}

				}

				break;
			}
		}

	}
}
