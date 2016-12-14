package agodahome;

import java.io.Console;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Homepage {
	static boolean play = true;
	static Scanner kb = new Scanner(System.in);

	public static void main(String[] args) throws IOException {

		System.out.println("Welcome to the Agoda Hotel");
		System.out
				.print("(1) Agoda employee \n(2) User Client \nPlease choose the number = ");
		int a = nextValidInt(kb);
		while (play) {

			if (a == 1) {
				emfunc();
			} else if (a == 2) {
				userfunc();
			} else
				System.out.println("Invalid number");

			System.out.print("Continue to use? (Y/N) = ");
			String con = kb.next();
			if (con.equalsIgnoreCase("n"))
				play = false;
			else if (con.equalsIgnoreCase("y")) {
				System.out.println();
			} else
				play = false;
		}
		System.out.print("Thank you!! :)");
		kb.close();
		System.exit(0);

	}

	public static void add(ArrayList<String> slist) throws IOException {

		File file = new File("src/agoda.xlsx");
		FileInputStream fIP = new FileInputStream(file);
		// Get the workbook instance for XLSX file
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);

		// Create a blank sheet
		XSSFSheet spreadsheet = workbook.getSheetAt(0);
		// Create row object
		XSSFRow row;
		// This data needs to be written (Object[])
		Map<String, Object[]> empinfo = new TreeMap<String, Object[]>();
		if (slist.size() < 8) {
			String s[] = new String[8];
			for (int i = 0; i < slist.size(); i++) {
				s[i] = slist.get(i);
			}
			empinfo.put("1", new Object[] { s[0], s[1], s[2], s[3], s[4], s[5],
					s[6], s[7] });

		} else
			empinfo.put(
					"1",
					new Object[] { slist.get(0), slist.get(1), slist.get(2),
							slist.get(3), slist.get(4), slist.get(5),
							slist.get(6), slist.get(7) });

		// Iterate over data and write to sheet
		Set<String> keyid = empinfo.keySet();

		int rowid = spreadsheet.getPhysicalNumberOfRows(); // spreadsheet.getLastRowNum()
															// + 1;
		// System.out.println(rowid);
		// ===set ID
		Row row1 = spreadsheet.getRow(rowid - 1);
		Cell cell1 = row1.getCell(0);

		String ids = Integer.parseInt(cell1.getStringCellValue()) + 1 + "";
		// System.out.println(ids + "test");
		// =====
		for (String key : keyid) {
			row = spreadsheet.createRow(rowid++);
			Object[] objectArr = empinfo.get(key);
			int cellid = 0;
			Cell cell = row.createCell(cellid++);
			cell.setCellValue(ids);
			for (Object obj : objectArr) {
				cell = row.createCell(cellid++);
				cell.setCellValue((String) obj);
			}
		}
		// Write the workbook in file system
		FileOutputStream out = new FileOutputStream(new File("src/agoda.xlsx"));
		workbook.write(out);
		out.close();
		workbook.close();
		System.out.println("Agoda Added successfully");

	}

	public static void removeId(String id) throws IOException {
		File file = new File("src/agoda.xlsx");
		FileInputStream fIP = new FileInputStream(file);
		// Get the workbook instance for XLSX file
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		XSSFSheet spreadsheet = workbook.getSheetAt(0);

		int rowremove = findRowId(spreadsheet, id);
		if (rowremove > -1) {
			removeRow(spreadsheet, rowremove);
			// System.out.println(rowremove);
		} else
			System.out.println("Not found");
		FileOutputStream out = new FileOutputStream(new File("src/agoda.xlsx"));
		workbook.write(out);
		out.close();
		workbook.close();
		System.out.println("Agoda Removed \"" + id + "\" successfully");

	}

	public static void updateId(String id) throws IOException {
		File file = new File("src/agoda.xlsx");
		FileInputStream fIP = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		XSSFSheet spreadsheet = workbook.getSheetAt(0);
		int rowid = findRowId(spreadsheet, id);
		// ====Print excel data
		if (rowid != -1) { // found id
			System.out
					.println("[ID]\t\t[Name]\t\t[Country]\t\t[City]\t\t[Street]"
							+ "\t\t[Telephone]\t\t[StarRanking]\t\t[Rooms]\t\t[Description]");
			printexcel(spreadsheet, rowid);
			System.out.print("\nWhat's index = ");
			int colindex = nextValidInt(kb);
			kb.nextLine();
			if (colindex == 0) {
				System.out.println("Can't change \"IDs\" ");
				workbook.close();
				return;
			} else if (colindex > 8 || colindex < 0) {
				System.out.println("Out of index");
				workbook.close();
				return;
			}
			System.out.print("To be = ");

			// ----- edit cell
			Row row = spreadsheet.getRow(rowid);
			Cell cell = row.getCell(colindex);
			cell.setCellValue(kb.nextLine());
			// -----
			FileOutputStream out = new FileOutputStream(new File(
					"src/agoda.xlsx"));
			workbook.write(out);
			out.close();
			workbook.close();
			System.out.println("Agoda Updated successfully");
			System.out
					.println("ID\t\tName\t\tCountry\t\tCity\t\tStreet\t\tTelephone\t\tStarRanking\t\tRooms\t\tDescription");
			printexcel(spreadsheet, rowid);

		} else
			System.out.println("Not found id = " + id);
		System.out.println();

		// ======

	}

	public static ArrayList<Integer> search(XSSFSheet sheet,
			ArrayList<String> items) {
		int a = 0;
		boolean b = false;
		ArrayList<Integer> myList = new ArrayList<Integer>();

		for (int i = 0; i < items.size(); i++) {
			if (i == 0) { // First time search
				for (Row row : sheet) {
					if (row.getRowNum() == 0) {
						//Skip Header
					} else {
						for (Cell cell : row) {
							cell.setCellType(Cell.CELL_TYPE_STRING); // HandleNumbericCell
							if (items
									.get(i)
									.toLowerCase()
									.equals(cell.getStringCellValue()
											.toLowerCase())
									|| (cell.getColumnIndex() == 1 && cell
											// for-name-search-contains-with-ignoreCase
											.getStringCellValue()
											.toLowerCase()
											.contains(
													items.get(i).toLowerCase())))
								myList.add(row.getRowNum());
						}
					}
				}
				// ====== for clear duplicate
				Set<Integer> hs = new HashSet<>();
				hs.addAll(myList);
				myList.clear();
				myList.addAll(hs);
				// ========
			} else { // Next Search
				a = 0;
				// System.out.println(myList);
				while (a < myList.size()) {
					for (Cell cell : sheet.getRow(myList.get(a))) {
						cell.setCellType(Cell.CELL_TYPE_STRING); // for
																	// handle-NUMBERIC-CELL
						if (items.get(i).equals(cell.getStringCellValue())
								|| (cell.getColumnIndex() == 1 && cell
										// for-name-search-contains-with-ignoreCase
										.getStringCellValue().toLowerCase()
										.contains(items.get(i).toLowerCase()))) {
							b = false;
						} else { // for-check-if-not-found
							// b = true;
						}
					}
					if (b) { // remove-if-not-found-in-the-next-search
						myList.remove(a);
						// a++;
						b = true;
					}
					a++;
					// System.out.println(myList);

				}
			}
		}
		return myList;
	}

	private static int findRowId(XSSFSheet sheet, String cellContent) { // notHandleIfid-name-same-as-others.
		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
					if (cell.getRichStringCellValue().getString().trim()
							.equals(cellContent)) {
						return row.getRowNum();
					}
				}
			}
		}
		return -1;
	}

	public static void removeRow(XSSFSheet sheet, int rowIndex) {
		int lastRowNum = sheet.getLastRowNum();
		if (rowIndex >= 0 && rowIndex < lastRowNum) {
			sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
		}
		if (rowIndex == lastRowNum) {
			XSSFRow removingRow = sheet.getRow(rowIndex);
			if (removingRow != null) {
				sheet.removeRow(removingRow);
			}
		}
	}

	public static void userfunc() throws IOException {
		File file = new File("src/agoda.xlsx");
		FileInputStream fIP = new FileInputStream(file);
		// Get the workbook instance for XLSX file
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);

		// Create a blank sheet
		XSSFSheet spreadsheet = workbook.getSheetAt(0);

		Scanner sc = new Scanner(System.in);
		System.out.println("What's the hotel do you want to search ? ");
		// String sfind = sc.nextLine();
		// List<String> items = Arrays.asList(sfind.split("\\s*,\\s*"));
		ArrayList<String> slist = new ArrayList<String>();
		System.out.print("Name = ");
		slist.add(sc.nextLine().trim());
		System.out.print("City = ");
		slist.add(sc.nextLine().trim());
		System.out.print("Star Ranking = ");
		slist.add(sc.nextLine().trim());
		System.out.println("Please wait...");
		ArrayList<Integer> myList = search(spreadsheet, slist);
		if (myList.size() > 0) {
			// ====Print excel data
			System.out
					.println("[ID]\t\t[Name]\t\t[Country]\t\t[City]\t\t[Street]"
							+ "\t\t[Telephone]\t\t[StarRanking]\t\t[Rooms]\t\t[Description]");
			for (int i = 0; i < myList.size(); i++) {
				printexcel(spreadsheet, myList.get(i));

				System.out.println();
			}
			// ======
		} else
			System.out.println("Sorry, Not found");
		workbook.close();
	}

	public static void emfunc() throws IOException {
		System.out.print("Please select the function you want to use \n"
				+ "(1) Add Hotel \n" + "(2) Remove Hotel \n"
				+ "(3) Update Hotel \n = ");
		Scanner sc = new Scanner(System.in);
		int b = nextValidInt(sc);
		sc.nextLine();
		if (b == 1) {
			System.out.println("Insert information for the hotel");
			ArrayList<String> slist = new ArrayList<String>();
			System.out.print("Name = ");
			slist.add(sc.nextLine().trim());
			System.out.print("Country = ");
			slist.add(sc.nextLine().trim());
			System.out.print("City = ");
			slist.add(sc.nextLine().trim());
			System.out.print("Street = ");
			slist.add(sc.nextLine().trim());
			System.out.print("Telephone = ");
			slist.add(sc.nextLine().trim());
			System.out.print("Star Ranking = ");
			slist.add(sc.nextLine().trim());
			System.out.print("Rooms = ");
			slist.add(sc.nextLine().trim());
			System.out.print("Description = ");
			slist.add(sc.nextLine().trim());
			add(slist);
			/*
			 * String sfind = sc.nextLine(); List<String> items =
			 * Arrays.asList(sfind.split("\\s*,\\s*")); add(items);
			 */
		} else if (b == 2) {
			System.out.print("Remove information for the hotel id? = ");
			String sfind = sc.nextLine();
			removeId(sfind);
		} else if (b == 3) {
			System.out.print("What's id you want to update? = ");
			String id = sc.next();
			updateId(id);

		} else
			System.out.println("Invalid number");
	}

	public static int nextValidInt(Scanner s) {
		while (!s.hasNextInt())
			System.out.println(s.next()
					+ " is not a valid number. Please type again:");
		return s.nextInt();
	}

	public static void printexcel(XSSFSheet sheet, int rowid) {
		Iterator<Cell> cellIterator = sheet.getRow(rowid).cellIterator();
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
				System.out.print(cell.getNumericCellValue() + " \t\t ");
				break;
			case Cell.CELL_TYPE_STRING:
				System.out.print(cell.getStringCellValue() + " \t\t ");
				break;
			}
		}

	}
}
