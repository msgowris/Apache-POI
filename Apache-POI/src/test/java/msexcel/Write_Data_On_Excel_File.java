package msexcel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Write_Data_On_Excel_File {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet s = wb.createSheet("Employee Details");

		Object emp_data[][] = { { "Employee ID", "Employee Name", "Employee Designation" },
				{ "1", "Gowri", "Test Engineer" }, { "2", "Karthick", "Project Engineer" },
				{ "3", "Madhan", "QA Engineer" }

		};

		int rows = emp_data.length;
		System.out.println(rows);
		int columns = emp_data[0].length;
		System.out.println(columns);

		for (int r = 0; r < rows; r++) {
			XSSFRow row = s.createRow(r);
			for (int c = 0; c < columns; c++) {
				XSSFCell cell = row.createCell(c);
				Object value = emp_data[r][c];

				if (value instanceof String) {
					cell.setCellValue((String) value);
				}
				if (value instanceof Integer) {
					cell.setCellValue((Integer) value);
				}

			}
		}

		File write_data_excel_file = new File(".\\Test_Data\\Write_Data.xlsx");
		FileOutputStream fos = new FileOutputStream(write_data_excel_file);
		wb.write(fos);

		fos.close();

	}
}
