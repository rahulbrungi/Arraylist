import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;

public class CollectionsHashSet {
	private static final String FILE_NAME = "/Users/rahulbrungi/Desktop/db/Emails.xlsx";

	public static void main(String args[]) {

		ArrayList<String> l = new ArrayList<String>();
		HashSet<String> hs = new HashSet<String>(); 

		try {

			FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
			Workbook workbook = new XSSFWorkbook(excelFile);
			Sheet datatypeSheet = workbook.getSheetAt(0);
			Iterator<Row> iterator = datatypeSheet.iterator();
			// cell class of apache get cell add it to list
			int i = 0;
			for (Row row : datatypeSheet) {
				if (i == 0) {
					i++;
					continue;
				}

				Cell cell = row.getCell(0);
				hs.add(cell.toString());
				l.add(cell.toString());
			}
			

			// Using while loop for printing
			
			  Iterator<String> itr=hs.iterator(); while(itr.hasNext()){
				  System.out.println(itr.next());
	 }
			 
			// using enhanced for loop for printing
			/*for (String s : hs) {
				System.out.println(s);
			}*/
			  System.out.println("\nthe array list size is " + l.size() );

			  System.out.println("the hash set size is " + hs.size() +".\t there are reduced elements in has set since there are no duplicate elements present");

			workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
}
