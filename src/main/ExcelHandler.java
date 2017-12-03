package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.OptionalDouble;

import javax.swing.JFileChooser;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelHandler
{
	private List<Sheet> sheets = new ArrayList<>();
	private int numberOfSheets;
	private Map<String, Integer> subtestPossibleScores = new HashMap<>();
	private Map<String, List<Double>> subtestScoreMap = new HashMap<>();

	public void initialize()
	{
		getExcelSheet();

		for (Sheet sheet : sheets)
		{
			parseSheet(sheet);
		}
	}

	public void getExcelSheet()
	{
		//find the excel file
		//TODO: need to change this back before prod
		//		JFileChooser chooser = new JFileChooser(System.getProperty("user.home"));
		JFileChooser chooser = new JFileChooser("C:\\Users\\Trevyn\\git\\hal9000\\src");
		int option = chooser.showOpenDialog(null);

		File file = null;

		if (option == JFileChooser.APPROVE_OPTION)
		{
			file = chooser.getSelectedFile();
		}

		FileInputStream inputStream = null;
		Workbook workbook = null;

		try
		{
			//reading said file
			inputStream = new FileInputStream(file);

			workbook = new XSSFWorkbook(inputStream);
			numberOfSheets = workbook.getNumberOfSheets();
			for (int i = 0; i < numberOfSheets; i++)
			{
				sheets.add(workbook.getSheetAt(i));
			}
		}
		catch (IOException e)
		{
			System.out.println("there was a problem yo, check dis out: " + e.getMessage());
		}
		finally
		{
			try
			{
				workbook.close();
				inputStream.close();
			}
			catch (IOException e)
			{
				System.out.println("could not close the workbook or inputStream");
				e.printStackTrace();
			}
		}
	}

	public void parseSheet(Sheet sheet)
	{
		Iterator<Row> iterator = sheet.iterator();

		String subtestName = "";
		List<Double> subtestScoreList = new ArrayList<>();

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
					String stringCellValue = cell.getStringCellValue();

					//get name of subtest
					if (!stringCellValue.isEmpty() && stringCellValue.contains("Standard"))
					{
						subtestName = stringCellValue;
						getSubtestTitleData(stringCellValue);
					}
					break;
				case NUMERIC:
					double doubleCellValue = cell.getNumericCellValue();

					int titleCellRow = cell.getRowIndex();
					int titleCellCol = cell.getColumnIndex();
					//get this range, for every cell with a number, avg them all and get a percentage
					//4,1     30,14

					int rowIndex = cell.getRowIndex();
					int colIndex = cell.getColumnIndex();
					if (rowIndex >= 4 && rowIndex <= 30 && colIndex >= 1 && colIndex <= 14)
					{
						subtestScoreList.add(doubleCellValue);
					}
					break;
				default:
					break;
				}
			}
		}
		subtestScoreMap.put(subtestName, subtestScoreList);

		printData();
	}

	private void getSubtestTitleData(String cellValue)
	{
		String subtestName = cellValue;

		int possibleIndex = subtestName.indexOf(" possible");
		String possibleScore = subtestName.substring(possibleIndex - 2, possibleIndex).trim();
		int intPossibleScore = Integer.parseInt(possibleScore);

		//store that bitch
		subtestPossibleScores.put(subtestName, intPossibleScore);
	}

	private void printData()
	{
		for (Entry<String, List<Double>> entry : subtestScoreMap.entrySet())
		{
			String subtestTitle = entry.getKey();
			List<Double> scores = entry.getValue();

			int possibleScore = subtestPossibleScores.get(subtestTitle);

			OptionalDouble averageOptional = scores.stream().mapToDouble(a -> a).average();

			double average = averageOptional.getAsDouble();
			double possibleScoreDouble = possibleScore;
			double averagePercentage = (average / possibleScoreDouble) * 100;
			DecimalFormat df2 = new DecimalFormat("##.##");
			averagePercentage = Double.valueOf(df2.format(averagePercentage));

			System.out.println(subtestTitle + "   -   " + averagePercentage + "%");
		}
	}
}