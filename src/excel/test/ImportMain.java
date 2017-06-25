package excel.test;

import java.io.IOException;
import java.util.List;

import excel.imports.ExcelToDBImporter;
import excel.imports.ExcelToDBTeilnahme;

public class ImportMain {

	public static void main(String[] args) {
		String path = "C:\\Users\\Jonas\\Desktop\\ExcelTest\\Test2.xlsx";

		try {
			List<ExcelToDBTeilnahme> teilnahmen = ExcelToDBImporter.importTeilnahmen(path);
			for (ExcelToDBTeilnahme teilnahme : teilnahmen) {
				System.out.println(teilnahme.getKlasse() + "\t" + teilnahme.getDisziplinID() + "\t" + teilnahme.getSchuelerID());
			}
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
