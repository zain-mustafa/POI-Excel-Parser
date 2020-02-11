package test.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;


class ReadExcelDemo
{

    public static void main(String[] args)
    {
        try
        {
            //Variables to keep count of iterators
            int rowNumber = 0;
            int columnNumber = 0;
            String currentKey = new String();

            // Variables to create list of HashMaps
            HashMap<String, String> lanPairs = new HashMap<String, String>();
            ArrayList<HashMap<String, String>> languageColumn = new ArrayList<HashMap<String, String>>();

            FileInputStream file = new FileInputStream(new File("C://Users/zain.mustafa/Documents/SubscriptionsTranslationsMaster.xlsx"));

            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext())
            {
                rowNumber++;
                System.out.println("We are at row: " + rowNumber);

                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();

                columnNumber = 0;
                while (cellIterator.hasNext())
                {
                    columnNumber++;
                    System.out.println("We are at column Number: " + columnNumber);

                    Cell cell = cellIterator.next();
                    //Check the cell type and format accordingly
                    switch (cell.getCellType())
                    {
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "t");
                            break;
                        case Cell.CELL_TYPE_STRING:
                            if (columnNumber > 1) {
                                lanPairs = new HashMap<String, String>();
                                if (rowNumber == 1) {
                                    lanPairs.put(currentKey, cell.getStringCellValue());
                                    languageColumn.add(lanPairs);
                                } else {

                                    languageColumn.get(columnNumber-2).put(currentKey, cell.getStringCellValue());

//                                    lanPairs.put(currentKey, cell.getStringCellValue());
//                                    languageColumn.add(columnNumber-1, lanPairs);
                                }
                            } else {
                                System.out.println("Setting the Key: " + cell.getStringCellValue());
                                currentKey = cell.getStringCellValue();
                            }
//                            System.out.print(cell.getStringCellValue() + "                   ");
                            break;
                    }
                }
                System.out.println("");
//                System.out.println(languageColumn);
                System.out.println(languageColumn.size());

                for(int i = 0; i < languageColumn.size(); i++) {
                    System.out.println(languageColumn.get(i).get("locale"));
                }
            }
            file.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}