import java.io.FileOutputStream;
import java.io.IOException;
import java.util.InputMismatchException;
import java.util.Locale;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.github.javafaker.service.FakeValuesService;
import com.github.javafaker.service.RandomService;
import io.codearte.jfairy.Fairy;
import io.codearte.jfairy.producer.person.Person;

public class GenerateData {
	
	static XSSFWorkbook workbook=new XSSFWorkbook();
	static XSSFSheet sheet=workbook.createSheet("Data");

	public static void main(String[] args) throws IOException, InputMismatchException {
		
		try {
		Scanner scn = new Scanner(System.in);
		System.out.println("Enter the number of data you are expecting");
		int Lenght = scn.nextInt();
		
		for(int i=1; i<=Lenght; i++){
		
		  Fairy fairy = Fairy.create(); 
		  Person person = fairy.person(); 
		  String Name=person.getFirstName(); 
		  	 //System.out.println(person.getEmail()); // barker@yahoo.com
			 //System.out.println(person.getTelephoneNumber());
			 
		  FakeValuesService faker = new FakeValuesService(
		  new Locale("en-US"), new RandomService());
		  String PAN=faker.regexify("[A-Z]{3}P[A-Z]{1}[0-9]{4}[A-Z]{1}"); //will return something like "6bJ1"
		  System.out.println(i + " " + Name +" "+ PAN);
				//faker.letterify("12??89"); //will return something like "12hZ89"
				//faker.numerify("ABC##EF"); //will return something like "ABC99EF"
				//faker.bothify("12??##ED"); //will return something like "12iL27ED"
		  
		  XSSFRow row=sheet.createRow(0);
			row.createCell(0).setCellValue("#");
		  	row.createCell(1).setCellValue("Name");
		  	row.createCell(2).setCellValue("PAN");
			
			  for(int j=i; j<=Lenght; j++) {
			 
			   XSSFRow row1=sheet.createRow(j);
			  row1.createCell(0).setCellValue(j);
			  row1.createCell(1).setCellValue(Name);
			  row1.createCell(2).setCellValue(PAN); 
			  }
			  	
}
		FileOutputStream fis=new FileOutputStream("C:\\Users\\Public\\Data.xlsx");
		System.out.println("Created data successfully stored in the following path C:\\Users\\Public\\Data.xlsx");	
		workbook.write(fis);
		fis.close();
		workbook.close();
	}	catch (InputMismatchException e) {
		System.out.println("Entered Data is not acceptable...!!"); }	
	}
}