package msexcel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_Data_From_Excel_File {

	public static void main(String[] args) throws IOException {

		String excel_file_path = ".\\Test_Data\\LoginData.xlsx";
		FileInputStream fis = new FileInputStream(excel_file_path);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet s = wb.getSheet("Sheet1");
		
		int rows_count = s.getLastRowNum();
		System.out.println(rows_count);
		int columns_count = s.getRow(1).getLastCellNum();
		System.out.println(columns_count);

		for (int r = 1; r <= rows_count; r++) {
			XSSFRow row = s.getRow(r);
			for (int c = 0; c < columns_count; c++) {
				XSSFCell cell_value = row.getCell(c);
				switch (cell_value.getCellType()) {
				case STRING:
					System.out.print(cell_value.getStringCellValue());
					break;
				case NUMERIC:
					System.out.print(cell_value.getNumericCellValue());
					break;
				}
				System.out.print(" | ");
			}
			System.out.println();

		}

	}

}
