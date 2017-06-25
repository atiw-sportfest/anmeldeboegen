package excel.exports;

public class DBToExcelDisziplin implements Comparable<DBToExcelDisziplin>{

	private boolean mannschaft;
	private boolean zusatzpunkt;
	private int min, max;
	private String name;
	private int id;
	
	public DBToExcelDisziplin(String name, int id, boolean mannschaft, boolean zusatzpunkt, int min, int max) {
		super();
		this.mannschaft = mannschaft;
		this.zusatzpunkt = zusatzpunkt;
		this.min = min;
		this.max = max;
		this.name = name;
		this.id = id;
	}
	public boolean isMannschaft() {
		return mannschaft;
	}
	public void setMannschaft(boolean mannschaft) {
		this.mannschaft = mannschaft;
	}
	public boolean isZusatzpunkt() {
		return zusatzpunkt;
	}
	public void setZusatzpunkt(boolean zusatzpunkt) {
		this.zusatzpunkt = zusatzpunkt;
	}
	public int getMin() {
		return min;
	}
	public void setMin(int min) {
		this.min = min;
	}
	public int getMax() {
		return max;
	}
	public void setMax(int max) {
		this.max = max;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public int getId() {
		return id;
	}
	public void setId(int id) {
		this.id = id;
	}
	
	private int getCompValue(){
		int d = 0;
		if(mannschaft){
			d=3;
		}else{
			if(zusatzpunkt){
				d = 1;
			}else{
				d = 2;
			}
		}
		return d;
	}
	
	@Override
	public int compareTo(DBToExcelDisziplin o) {
		return Integer.compare(getCompValue(), o.getCompValue());
	}
	
}
