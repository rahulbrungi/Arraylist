mport org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;


public class CollectionsAssignment {

	private static final String FILE_NAME = "/Users/rahulbrungi/Desktop/db/Emails.xlsx";

	public static void main(String[] args) {
		ArrayList<String> l = new ArrayList<String>();

		try {

			FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
			Workbook workbook = new XSSFWorkbook(excelFile);
			Sheet datatypeSheet = workbook.getSheetAt(0);
			Iterator<Row> iterator = datatypeSheet.iterator();
			//cell class of apache get cell add it to list
int i=0;
			for(Row row : datatypeSheet) {
				if(i==0) {
				i++; continue;  }
				
Cell cell = row.getCell(0);
	l.add(cell.toString())	;
}
//System.out.println(l.size());

/*
 Iterator<String> itr=l.iterator();  
while(itr.hasNext()){  
 System.out.println(itr.next());  
}
*/

for (String s : l) {
	System.out.println(s);
}
			
			
			
			workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
System.out.print(l.size());
		
	}
}
