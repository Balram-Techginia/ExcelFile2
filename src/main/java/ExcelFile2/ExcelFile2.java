package ExcelFile2;


import com.mongodb.MongoClient;
import com.mongodb.MongoClientURI;
import com.mongodb.bulk.BulkWriteUpsert;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.model.DBCollectionUpdateOptions;
import com.mongodb.client.model.UpdateOptions;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bson.Document;
import java.util.Iterator;
import java.io.*;

import java.util.*;

public class ExcelFile2 {

	public static MongoClient getConnection() {
		try {
			return new MongoClient(new MongoClientURI("mongodb://admin:myadminpassword@ENTER your database ip"));
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}

	public static void main(String args[]) throws IOException {
		try {
			MongoClient client = getConnection();
			MongoCollection<Document> userCollection = client.getDatabase("QuickLookDB").getCollection("nikhil1");

			List<Document> Records = new ArrayList<Document>();
		
			
		File folder = new File("C:\\\\Users\\\\BALRAM\\\\Downloads\\ExcelFiles");
		File[] listOfFiles = folder.listFiles();

		for (File file : listOfFiles) {
			    if (file.isFile()) {
			        System.out.println(file.getName());		
			FileInputStream fis = new FileInputStream("DATA1");
			XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
		//	SXSSFWorkbook SWorkBook = new SXSSFWorkbook(myWorkBook);
			// SXSSFWorkbook wb = new SXSSFWorkbook(new XSSFWorkbook(), 100, true, true);
			Sheet mySheet = myWorkBook.getSheetAt(0);
			System.out.println(mySheet.getSheetName());
			String headerArr[] = new String[17];

			Iterator<Row> rowIterator = mySheet.iterator();
			Row headerRow = rowIterator.next();
			Iterator<Cell> headerCellIterator = headerRow.cellIterator();
			int i = 0;
			int count = 0;
			while (headerCellIterator.hasNext()) {
				Cell headerCell = headerCellIterator.next();
				headerArr[i] = headerCell.toString();
				i++;
			}
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Document rec = new Document();
				Document mob = new Document();
				i = 0; //int count=0;

				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					try {

						Cell cell = cellIterator.next();

						switch (cell.getCellType()) {

						case STRING:
							if(headerArr[i].equalsIgnoreCase("Mobile"))
							{
								mob.put(headerArr[i], cell.getStringCellValue());
									}
								else
								{
								rec.put(headerArr[i], cell.getStringCellValue());
							}
							
							//System.out.print(cell.getStringCellValue() + "\t\t\t");
							break;
						case NUMERIC:
							if(headerArr[i].equalsIgnoreCase("Mobile"))
							{
								mob.put(headerArr[i], cell.getStringCellValue());
									}
								else
								{
									rec.put(headerArr[i], (int) cell.getNumericCellValue());
							}
							
							break;

						case BLANK:
							rec.put(headerArr[i], "");
							break;
						default:
						}
						i++;
					} catch (Exception e) {
					}
				}

				userCollection.updateOne(mob, rec, "$set:"+ new UpdateOptions().upsert(true));							
					Records.clear();
			
				// records size is big
			}

			myWorkBook.close();
		//	SWorkBook.close();
			    } 
			   
			}

		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
