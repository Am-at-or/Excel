import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Main {

	public static void main(String[] args) {
		try {
			writeIntoExcel("D://Excel.xls");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		try {
			readFromExcel("D://Excel.xls");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	@SuppressWarnings("deprecation")
	public static void writeIntoExcel(String file)
			throws FileNotFoundException, IOException {
		Workbook book = new HSSFWorkbook();
		Sheet sheet = book.createSheet("Birthdays");

		// Нумерація починается з нуля
		Row row = sheet.createRow(0);

		// Записуємо ім'я і дату в два стовпця
		// ім'я буде String, а дата народження - Date,
		// формату dd.mm.yyyy
		Cell name = row.createCell(0);
		name.setCellValue("John");

		Cell birthdate = row.createCell(1);

		DataFormat format = book.createDataFormat();
		CellStyle dateStyle = book.createCellStyle();
		dateStyle.setDataFormat(format.getFormat("dd.mm.yyyy"));
		birthdate.setCellStyle(dateStyle);

		// Нумерація років починается с 1900-го
		birthdate.setCellValue(new Date(110, 10, 10));

		// Змінюємо размір стовпця
		sheet.autoSizeColumn(1);

		// Записуємо все в файл
		book.write(new FileOutputStream(file));
		book.close();
	}

	public static void readFromExcel(String file) throws IOException {
		HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(file));
		HSSFSheet myExcelSheet = myExcelBook.getSheet("Birthdays");
		HSSFRow row = myExcelSheet.getRow(0);

		if (row.getCell(0).getCellType() == HSSFCell.CELL_TYPE_STRING) {
			String name = row.getCell(0).getStringCellValue();
			System.out.println("name : " + name);
		}

		if (row.getCell(1).getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
			Date birthdate = row.getCell(1).getDateCellValue();
			System.out.println("birthdate :" + birthdate);
		}

		myExcelBook.close();

	}

}
