import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel4 {

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
			int rowCount = sheet.getLastRowNum();// Starts with index of 0
			int cellCount = sheet.getRow(0).getLastCellNum();// Gives the exact value

			Iterator<Row> rowIterator = sheet.iterator();
			Row roww = rowIterator.next();
			int colcountt = roww.getLastCellNum();
			int rowcnt = sheet.getLastRowNum();

			System.out.println("column count:: " + colcountt);
			System.out.println("rowcount :::" + rowcnt);

			outerloop: for (int i = 0; i < rowcnt; i++) {
				
				Row row = rowIterator.next();

				// For each row, iterate through all the columns
//	            Iterator<Cell> cellIterator = row.cellIterator();

				long gsmno = 0;
				String simno = "";
				String customerName = "";
//	            int simType = 0;
				int j = 0;

				int colcount = row.getLastCellNum();

				for (j = 0; j < colcount; j++) {
				
					Cell cell = row.getCell(j);

					System.out.println("cell::" + cell);

					/*
					 * if (i >= quantity) { break outerloop; }
					 */
					
					if (cell == null) {
						if (j == 0)
							returnmsg = "Please upload another file as it consists of empty GSM number at row "
									+ (i + 2);
						else if (j == 1)
							returnmsg = "Please upload another file as it consists of empty SIM number at row "
									+ (i + 2);

						validateObjResp.put("Message", returnmsg);
						validateArrResp.add(validateObjResp);
						return validateArrResp;
					}

					
					//Cell cell = cellIterator.next();
					// Check the cell type and format accordingly
					if (j == 0) {
						CRSAppLogger.debug("cell type gsm:: " + cell.getCellType());
						if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							if (((int) Math.floor(Math.log10((long) cell.getNumericCellValue()) + 1)) == 10
									|| ((int) Math.floor(Math.log10((long) cell.getNumericCellValue()) + 1)) == 13) {
								gsmno = (long) cell.getNumericCellValue();
							} else {
								returnmsg = "GSM number length should be 10 or 13 digits.Invalid GSM number at row "
										+ (i + 2);
								validateObjResp.put("Message", returnmsg);
								validateArrResp.add(validateObjResp);
								return validateArrResp;
							}
						} else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
							CRSAppLogger.debug("gsmno str" + cell.getStringCellValue());
							if ((cell.getStringCellValue()).length() > 0) {
								if ((cell.getStringCellValue()).length() == 10
										|| (cell.getStringCellValue()).length() == 13) {// allowing 13 digits mobile no
																						// also
									gsmno = Long.parseLong(cell.getStringCellValue());
								} else {
									returnmsg = "GSM number length should be 10 or 13 digits.Invalid GSM number at row "
											+ (i + 2);
									validateObjResp.put("Message", returnmsg);
									validateArrResp.add(validateObjResp);
									return validateArrResp;
								}
							} else {
								returnmsg = "Please upload another file as it consists of empty GSM number at row "
										+ (i + 2);
								validateObjResp.put("Message", returnmsg);
								validateArrResp.add(validateObjResp);
								return validateArrResp;
							}
						} else {
							returnmsg = "Please upload another file as it consists of invalid GSM number at row "
									+ (i + 2);
							validateObjResp.put("Message", returnmsg);
							validateArrResp.add(validateObjResp);
							return validateArrResp;
						}
					} else if (j == 1) {
						if (colcount < 2) {
							returnmsg = "Invalid file as uploaded file has less than three columns at row " + (i + 2)
									+ " ";
							validateObjResp.put("Message", returnmsg);
							validateArrResp.add(validateObjResp);
							return validateArrResp;
						}
						if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
							if ((cell.getStringCellValue()).length() == 19) {
								simno = cell.getStringCellValue();
							} else {
								returnmsg = "Please upload another file as it consists of invalid SIM number at row "
										+ (i + 2);
								validateObjResp.put("Message", returnmsg);
								validateArrResp.add(validateObjResp);
								return validateArrResp;
							}
						} else {
							returnmsg = "Please upload another file as it consists of invalid SIM number at row "
									+ (i + 2);
							validateObjResp.put("Message", returnmsg);
							validateArrResp.add(validateObjResp);
							return validateArrResp;
						}
					} else if (j == 2) {// customer name column optional

						if (colcount > 3) {
							returnmsg = "Invalid file as uploaded file has more than three columns at row " + (i + 2)
									+ " ";
							validateObjResp.put("Message", returnmsg);
							validateArrResp.add(validateObjResp);
							return validateArrResp;
						}
						if (colcount == 3) {
							if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
								Cell custCell = row.getCell(j);
								if (custCell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) {
									customerName = (cell.getStringCellValue());
								}

							}
						}
					} else {
						j = 0;
					}
				}
//	            resJson.put("SimType", simType);	
				if (gsmno == 0 || String.valueOf(gsmno).length() == 0) {
					returnmsg = "Please upload another file as it consists of GSM number as empty or 0 at row "
							+ (i + 2);
					validateObjResp.put("Message", returnmsg);
					validateArrResp.add(validateObjResp);
					return validateArrResp;
				} else {
					if (simno.equalsIgnoreCase("0") || simno.length() == 0) {
						returnmsg = "Please upload another file as it consists of SIM number as empty or 0 at row "
								+ (i + 2);
						validateObjResp.put("Message", returnmsg);
						validateArrResp.add(validateObjResp);
						return validateArrResp;
					} else {
						resJson = new JSONObject();
						resJson.put(gsmno, simno);

						if (customerName.length() > 0) {
							resJson.put("Customer_name", customerName);
						} else {
							resJson.put("Customer_name", "");
						}
						resultJson.add(resJson); // storing excel file data into jsonarray
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
