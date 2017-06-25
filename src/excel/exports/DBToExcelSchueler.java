package excel.exports;

public class DBToExcelSchueler {

	private String name, vorname;
	private int id;
	
	public DBToExcelSchueler(String name, String vorname, int id) {
		super();
		this.name = name;
		this.vorname = vorname;
		this.id = id;
	}
	
	
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getVorname() {
		return vorname;
	}
	public void setVorname(String vorname) {
		this.vorname = vorname;
	}
	public int getId() {
		return id;
	}
	public void setId(int id) {
		this.id = id;
	}
	
	
	
}
