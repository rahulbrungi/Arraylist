import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Comparator;
import java.util.TreeSet;;

public class CollectionTreeSet {
	private static final String FILE_NAME = "/Users/rahulbrungi/Desktop/db/person.xlsx";

	public static void main(String args[]) throws IOException {

		TreeSet<PersonDetails> pd = new TreeSet<PersonDetails>(new Logic());

		FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
		Workbook workbook = new XSSFWorkbook(excelFile);
		Sheet dataTypeSheet = workbook.getSheetAt(0);

		Row row = dataTypeSheet.getRow(1);
		PersonDetails p1 = new PersonDetails(row.getCell(0).toString(), row.getCell(1).toString(),
				row.getCell(2).toString());
		pd.add(p1);

		row = dataTypeSheet.getRow(2);
		PersonDetails p2 = new PersonDetails(row.getCell(0).toString(), row.getCell(1).toString(),
				row.getCell(2).toString());
		pd.add(p2);

		row = dataTypeSheet.getRow(3);
		PersonDetails p3 = new PersonDetails(row.getCell(0).toString(), row.getCell(1).toString(),
				row.getCell(2).toString());
		pd.add(p3);

		row = dataTypeSheet.getRow(4);
		PersonDetails p4 = new PersonDetails(row.getCell(0).toString(), row.getCell(1).toString(),
				row.getCell(2).toString());
		pd.add(p4);

		row = dataTypeSheet.getRow(5);
		PersonDetails p5 = new PersonDetails(row.getCell(0).toString(), row.getCell(1).toString(),
				row.getCell(2).toString());
		pd.add(p5);

		for (PersonDetails p : pd) {
			System.out.println(p.id + " " + p.name + " " + p.description);
		}
		workbook.close();

	}
}

class PersonDetails {

	String id;
	String name;
	String description;

	public PersonDetails(String id, String name, String description) {
		this.id = id;
		this.name = name;
		this.description = description;

	}

	public String getName() {
		return name;
	}

	@Override
	public boolean equals(Object obj) {
		// TODO Auto-generated method stub
		return super.equals(obj);
	}

}

class Logic implements Comparator<PersonDetails> {

	public int compare(PersonDetails p1, PersonDetails p2) {
		return p1.getName().compareTo(p2.getName());
	}
}
