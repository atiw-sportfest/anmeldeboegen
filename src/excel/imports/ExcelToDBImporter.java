package excel.imports;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToDBImporter {

	public static List<ExcelToDBTeilnahme> importTeilnahmen(String path) throws IOException{
		return importTeilnahmen(new FileInputStream(path));
	}
	
	public static List<ExcelToDBTeilnahme> importTeilnahmen(InputStream inputStream) throws IOException{
		List<ExcelToDBTeilnahme> teilnahmen = new ArrayList<>();
		
		XSSFSheet sheet = new XSSFWorkbook(inputStream).getSheet("Anmeldebogen");
		
		String klasse = sheet.getRow(0).getCell(1).getStringCellValue();
		int schuelerCount = (int) sheet.getRow(1).getCell(0).getNumericCellValue();
		int disziplinCount = (int) sheet.getRow(2).getCell(0).getNumericCellValue();
		
		for(int i = 0; i < schuelerCount; i++){
			for(int j = 0; j < disziplinCount; j++){
				if(!sheet.getRow(i+4).getCell(j+3).getStringCellValue().isEmpty()){
					teilnahmen.add(new ExcelToDBTeilnahme((int)sheet.getRow(i+4).getCell(0).getNumericCellValue(), (int)sheet.getRow(3).getCell(j+3).getNumericCellValue(), klasse));
				}
			}
		}
		
		return teilnahmen;
	}
	
}
