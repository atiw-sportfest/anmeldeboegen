package excel.imports;

public class ExcelToDBTeilnahme {

	private int schuelerID, disziplinID;
	private String klasse;
	
	public ExcelToDBTeilnahme(int schuelerID, int disziplinID, String klasse) {
		this.schuelerID = schuelerID;
		this.disziplinID = disziplinID;
		this.klasse = klasse;
	}

	public int getSchuelerID() {
		return this.schuelerID;
	}

	public void setSchuelerID(int schuelerID) {
		this.schuelerID = schuelerID;
	}

	public int getDisziplinID() {
		return this.disziplinID;
	}

	public void setDisziplinID(int disziplinID) {
		this.disziplinID = disziplinID;
	}

	public String getKlasse() {
		return this.klasse;
	}

	public void setKlasse(String klasse) {
		this.klasse = klasse;
	}	
	
}


