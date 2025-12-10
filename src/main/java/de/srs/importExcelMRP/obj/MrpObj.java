package de.srs.importExcelMRP.obj;

import java.sql.Timestamp;

public class MrpObj {
	
	private String KosTr;
	private String KosTr_Bez;
	private String Kostr_Grp;
	private String Frmt;
	private Timestamp Datum_Zeit;


	public String getKosTr() {
		return KosTr;
	}
	public void setKosTr(String kosTr) {
		KosTr = kosTr;
	}
	public String getKosTr_Bez() {
		return KosTr_Bez;
	}
	public void setKosTr_Bez(String kosTr_Bez) {
		KosTr_Bez = kosTr_Bez;
	}
	public String getKostr_Grp() {
		return Kostr_Grp;
	}
	public void setKostr_Grp(String kostr_Grp) {
		Kostr_Grp = kostr_Grp;
	}
	public Timestamp getDatum_Zeit() {
		return Datum_Zeit;
	}
	public String getFrmt() {
		return Frmt;
	}
	public void setFrmt(String frmt) {
		Frmt = frmt;
	}	
	
	public void setDatum_Zeit(Timestamp datum_Zeit) {
		Datum_Zeit = datum_Zeit;
	}
	
	public MrpObj() {
		super();
	}
	
	public MrpObj(String kosTr, String kosTr_Bez, String kostr_Grp, String frmt, Timestamp datum_Zeit) {
		super();
		KosTr = kosTr;
		KosTr_Bez = kosTr_Bez;
		Kostr_Grp = kostr_Grp;
		Frmt = frmt;
		Datum_Zeit = datum_Zeit;
	}
	
	  @Override
	    public String toString() {
	        return " Mrp{KosTr=" + KosTr + ", KosTr_Bez=" + KosTr_Bez +  " Kostr_Grp=" + Kostr_Grp + " Frmt=" + Frmt + " Datum_Zeit="+ Datum_Zeit + "}";
	    }

}
