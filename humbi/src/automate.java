import java.io.File;  
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.Row; 
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  

class automate
{  
	public static void main(String[] args)   
	{  
		try  
		{  
			String path = "C:\\Users\\ishra\\eclipse-workspace\\humbi\\src\\InputDataSet.xlsx";
			File file = new File(path);   //creating a new file instance  
			FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
			//creating Workbook instance that refers to .xlsx file  
			XSSFWorkbook wb = new XSSFWorkbook(fis);   
			XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object  
			Iterator<Row> itr = sheet.iterator();    //iterating over excel file  
			Workbook workbook = new HSSFWorkbook();
			Sheet opsheet = workbook.createSheet("Sheet1");
			
			int i = 0;
			Row oprow = opsheet.createRow(i);
			Cell opcell = oprow.createCell(0);
			
			
			while (itr.hasNext())                 
			{  
				Row row = itr.next();  
				Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column  
				while (cellIterator.hasNext())   
				{  
					Cell cell = cellIterator.next();  
					switch (cell.getCellType())               
					{  
					case Cell.CELL_TYPE_STRING:    //field that represents string cell type  
						System.out.print(cell.getStringCellValue() + "\t\t\t");
						opcell.setCellValue(cell.getStringCellValue());
						oprow = opsheet.createRow(i++);
						opcell = oprow.createCell(0);
						break;  
					case Cell.CELL_TYPE_NUMERIC:    //field that represents number cell type  
						System.out.print(cell.getNumericCellValue() + "\t\t\t");
						opcell.setCellValue(cell.getNumericCellValue());
						oprow = opsheet.createRow(i++);
						opcell = oprow.createCell(0);
						break;  
					default:  
					}  
				}  
				System.out.println("");  
			}
			FileOutputStream fos = new FileOutputStream("C:\\Users\\ishra\\eclipse-workspace\\humbi\\src\\Output.xlsx");
	        workbook.write(fos);
	        fos.close();
		}  
	catch(Exception e)  
		{  
			e.printStackTrace();  
		}  
	}  
}  
