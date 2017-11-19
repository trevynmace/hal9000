package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import javax.swing.JFileChooser;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader
{
	public static void main(String[] args)
	{
		JFileChooser chooser = new JFileChooser(System.getProperty("user.home"));
		int option = chooser.showOpenDialog(null);

		File file = null;

		if (option == JFileChooser.APPROVE_OPTION)
		{
			file = chooser.getSelectedFile();
		}

		try
		{
			FileInputStream inputStream = new FileInputStream(file);

			Workbook workbook = new XSSFWorkbook(inputStream);
			Sheet firstSheet = workbook.getSheetAt(0);
			Iterator<Row> iterator = firstSheet.iterator();

			while (iterator.hasNext())
			{
				Row nextRow = iterator.next();
				Iterator<Cell> cellIterator = nextRow.cellIterator();

				while (cellIterator.hasNext())
				{
					Cell cell = cellIterator.next();

					switch (cell.getCellTypeEnum())
					{
					case STRING:
						System.out.print(cell.getStringCellValue());
						break;
					case BOOLEAN:
						System.out.print(cell.getBooleanCellValue());
						break;
					case NUMERIC:
						System.out.print(cell.getNumericCellValue());
						break;
					default:
						break;
					}
					System.out.print(" - ");
				}
				System.out.println();
			}

			workbook.close();
			inputStream.close();
		}
		catch (IOException e)
		{
			System.out.println("there was a problem yo, check dis out: " + e.getMessage());
		}
	}
}