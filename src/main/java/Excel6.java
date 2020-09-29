import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel6 {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		ArrayList<String> finalOutput = getData("testdata", "TestCases", "AddProfile");

		int size = finalOutput.size();

		for (int i = 1; i < size; i++) {

			System.out.println(finalOutput.get(i));
		}

		System.out.println("===============================================================");

		ArrayList<String> finalOutput2 = getData("testdata", "TestCases", "DeleteProfile");

		int size2 = finalOutput2.size();

		for (int i = 1; i < size2; i++) {

			System.out.println(finalOutput2.get(i));
		}

		System.out.println("===============================================================");

		ArrayList<String> finalOutput3 = getData("negativedata", "TestData", "isinNumber");

		int size3 = finalOutput3.size();

		for (int i = 1; i < size3; i++) {

			System.out.println(finalOutput3.get(i));
		}

		System.out.println("===============================================================");

		ArrayList<String> finalOutput4 = getData("negativedata", "TestData", "sharename");

		int size4 = finalOutput4.size();

		for (int i = 1; i < size4; i++) {

			System.out.println(finalOutput4.get(i));
		}

	}

	public static ArrayList<String> getData(String sheetName, String testCaseColumnName, String testcaseName)
			throws IOException {

		// fileInputStream argument
		ArrayList<String> a = new ArrayList<String>();

		FileInputStream inputFile = new FileInputStream("C://Users//aravindkoduri//Desktop//Test.xlsx");
		XSSFWorkbook fileWorkBook = new XSSFWorkbook(inputFile);

		int totalNumberOfSheets = fileWorkBook.getNumberOfSheets();

		for (int i = 0; i < totalNumberOfSheets; i++) {

			if (fileWorkBook.getSheetName(i).equalsIgnoreCase(sheetName)) {

				XSSFSheet targetSheet = fileWorkBook.getSheetAt(i);

				// Identify Testcases coloum by scanning the entire 1st row
				Iterator<Row> totalSheetRows = targetSheet.iterator();// sheet is collection of rows
				Row firstrow = totalSheetRows.next();

				Iterator<Cell> cellInTheRow = firstrow.cellIterator();// row is collection of cells
				int k = 0;
				int coloumn = 0;

				while (cellInTheRow.hasNext()) {

					Cell respectiveTestCaseColumnName = cellInTheRow.next();

					if (respectiveTestCaseColumnName.getStringCellValue().equalsIgnoreCase(testCaseColumnName)) {

						coloumn = k;

						System.out.println("Column value ::: " + coloumn);

						System.out.println("K vaue after in loop " + k);
					}

					k++;

					System.out.println("K vaue after inc " + k);
				}

				/* System.out.println(coloumn); */ // This line to print the column matching number.

				while (totalSheetRows.hasNext()) {

					Row targetedRow = totalSheetRows.next();

					if (targetedRow.getCell(coloumn).getStringCellValue().equalsIgnoreCase(testcaseName)) {

						Iterator<Cell> targetedCellValue = targetedRow.cellIterator();

						while (targetedCellValue.hasNext()) {

							Cell targetedCell = targetedCellValue.next();

							a.add(targetedCell.getStringCellValue());

						}
					}

				}

			}
		}
		return a;

	}

}
