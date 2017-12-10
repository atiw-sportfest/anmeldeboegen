package excel.test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import excel.exports.DBToExcelDisziplin;
import excel.exports.DBToExcelExporter;
import excel.exports.DBToExcelSchueler;

public class ExportMain {

	public static void main(String[] args) {
		List<DBToExcelSchueler> schueler = new ArrayList<DBToExcelSchueler>();
		schueler.add(new DBToExcelSchueler("Baehr", "Steffen", 1));
		schueler.add(new DBToExcelSchueler("Boeckmann", "Mirco", 2));
		schueler.add(new DBToExcelSchueler("Braun", "Alexander", 3));
		schueler.add(new DBToExcelSchueler("Brinkmann", "Maja", 4));
		schueler.add(new DBToExcelSchueler("Daum", "Tobias", 5));
		schueler.add(new DBToExcelSchueler("Deselaers", "Robin", 6));
		schueler.add(new DBToExcelSchueler("Doerper", "Tobias", 7));
		schueler.add(new DBToExcelSchueler("Dunsche", "Michael", 8));
		schueler.add(new DBToExcelSchueler("Eisenhofer", "Florian", 9));
		schueler.add(new DBToExcelSchueler("Hammerschmidt", "Jonas", 10));
		schueler.add(new DBToExcelSchueler("Hartleitner", "Sebastian", 11));
		schueler.add(new DBToExcelSchueler("Meinl", "Christian", 12));
		schueler.add(new DBToExcelSchueler("Neiss", "Benjamin", 13));
		schueler.add(new DBToExcelSchueler("Pauli", "David Florian", 14));
		schueler.add(new DBToExcelSchueler("Rieskamp", "Jonas", 15));
		schueler.add(new DBToExcelSchueler("Sawyerr", "David", 16));
		schueler.add(new DBToExcelSchueler("Schpke", "Lars", 17));
		schueler.add(new DBToExcelSchueler("Schroeder", "Maximilian", 18));
		schueler.add(new DBToExcelSchueler("Schulz", "Christoph", 19));
		schueler.add(new DBToExcelSchueler("Wennemaring", "David", 20));
		
		List<DBToExcelDisziplin> disziplinen = new ArrayList<DBToExcelDisziplin>();
		disziplinen.add(new DBToExcelDisziplin("2000m Lauf", 1, false, 3, 3));
		disziplinen.add(new DBToExcelDisziplin("4x100m Staffel", 2, false, 4, 4));
		disziplinen.add(new DBToExcelDisziplin("Baseball", 3, false, 5, 5));
		disziplinen.add(new DBToExcelDisziplin("Basketball", 4, false, 5, 5));
		disziplinen.add(new DBToExcelDisziplin("Beach - Volleyball", 5, true, 3, 3));
		disziplinen.add(new DBToExcelDisziplin("Frisbee", 6, false, 5, 5));
		disziplinen.add(new DBToExcelDisziplin("Fußball", 7, true, 5, 5));
		disziplinen.add(new DBToExcelDisziplin("Hochstrung", 8, false, 3, 3));
		disziplinen.add(new DBToExcelDisziplin("Hockey", 9, true, 4, 4));
		disziplinen.add(new DBToExcelDisziplin("Kisten waagerecht", 10, true, 5, 5));
		disziplinen.add(new DBToExcelDisziplin("Medizinballweitwurf", 11, false, 5, 5));
		disziplinen.add(new DBToExcelDisziplin("Weitsprung", 12, false, 5, 5));
				
		String klasse = "FS151";
		
		try {
			DBToExcelExporter.export(new FileOutputStream("ExportText.xsls"), klasse, schueler, disziplinen);
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}

}
