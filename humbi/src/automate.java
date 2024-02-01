import java.io.File;  
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

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
			String path = "C:\\Users\\mohammed.ishrath\\git\\Automate\\humbi\\src\\InputDataSet.xlsx";
			File file = new File(path);   //creating a new file instance  
			FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
			//creating Workbook instance that refers to .xlsx file  
			XSSFWorkbook wb = new XSSFWorkbook(fis);   
			XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object  
			Iterator<Row> itr = sheet.iterator();    //iterating over excel file  
			Workbook workbook = new HSSFWorkbook();

			List<List<Integer>> list = new ArrayList<>();
			
			
			while (itr.hasNext())                 
			{  
				List<Integer> subList = new ArrayList<>();
				Row row = itr.next();  
				Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column  
				while (cellIterator.hasNext())   
				{  
					Cell cell = cellIterator.next();  
					switch (cell.getCellType())               
					{  
					case Cell.CELL_TYPE_STRING:
						break;  
					case Cell.CELL_TYPE_NUMERIC:    
						int x = (int) cell.getNumericCellValue();
						subList.add(x);
						break;  
					default:  
					}  
				}  
				list.add(subList);
			}
	        
	        
	        Sheet opsheet = workbook.createSheet("Sheet1");		
			int i = 0;
			Row oprow = opsheet.createRow(i);
			Cell opcell = oprow.createCell(0);
	        int count=1;
	        int listLen = list.size();
	        int j = 0;
	        int subListIndex = 0;
	        
	        opcell.setCellValue(list.get(0).get(0));
			oprow = opsheet.createRow(i++);
			opcell = oprow.createCell(0);
	        
	        while(subListIndex < list.get(0).size()) {
	        	for(int lindex = 0; lindex < list.size(); lindex++) {
	        		
	        		List<Integer> currentSublist = list.get(lindex);
	                if (subListIndex < currentSublist.size()) {
	                	System.out.println(list.get(lindex).get(subListIndex)+" "+count++);
			        	opcell.setCellValue(list.get(lindex).get(subListIndex));
						oprow = opsheet.createRow(i++);
						opcell = oprow.createCell(0);
	                }
	                else {
	                	break;
	                }
		        }
	        	subListIndex++;
	        }
	        FileOutputStream fos = new FileOutputStream("C:\\Users\\mohammed.ishrath\\git\\Automate\\humbi\\src\\Output.xls");
	        workbook.write(fos);
	        fos.close();
		}  
	catch(Exception e)  
		{  
			e.printStackTrace();  
		}  
	}  
}  
