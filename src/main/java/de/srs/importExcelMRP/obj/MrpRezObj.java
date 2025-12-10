package de.srs.importExcelMRP.obj;

public class MrpRezObj {
	
	private String KosTr;
	private String OfenTyp;
	private int Zeile;
	private String MatNr;
	private String MatGrp;
	private String MatName;
	private Float Rezeptanteil;
	private Float Preis;
	private Float Ausbeute;
	private int Monat;
    private String GattGrp;

    public MrpRezObj() {
		super();
	}

	public MrpRezObj(String kosTr, String OfenTyp, int zeile, String matNr, String matGrp, String matName, Float rezeptanteil,
			Float preis, Float ausbeute, int monat) {
		super();
		KosTr = kosTr;
		OfenTyp = OfenTyp;
		Zeile = zeile;
		MatNr = matNr;
		MatGrp = matGrp;
		MatName = matName;
		Rezeptanteil = rezeptanteil;
		Preis = preis;
		Ausbeute = ausbeute;
		Monat = monat;
	}

    public void setOfenTyp(String ofenTyp) {
        OfenTyp = ofenTyp;
    }

    public String getKosTr() {
		return KosTr;
	}

	public void setKosTr(String kosTr) {
		KosTr = kosTr;
	}

	public String getOfenTyp() {
		return OfenTyp;
	}
	public int getZeile() {
		return Zeile;
	}

	public void setZeile(int zeile) {
		Zeile = zeile;
	}

	public String getMatNr() {
		return MatNr;
	}

	public void setMatNr(String matNr) {
		MatNr = matNr;
	}

	public String getMatGrp() {
		return MatGrp;
	}

	public void setMatGrp(String matGrp) {
		MatGrp = matGrp;
	}

	public String getMatName() {
		return MatName;
	}

	public void setMatName(String matName) {
		MatName = matName;
	}

	public Float getRezeptanteil() {
		return Rezeptanteil;
	}

	public void setRezeptanteil(Float rezeptanteil) {
		Rezeptanteil = rezeptanteil;
	}

	public Float getPreis() {
		return Preis;
	}

	public void setPreis(Float preis) {
		Preis = preis;
	}

	public Float getAusbeute() {
		return Ausbeute;
	}

	public void setAusbeute(Float ausbeute) {
		Ausbeute = ausbeute;
	}

	public int getMonat() {
		return Monat;
	}

	public void setMonat(int monat) {
		Monat = monat;
	}

    public String getGattGrp() {
        return GattGrp;
    }

    public void setGattGrp(String gattGrp) {
        GattGrp = gattGrp;
    }

    @Override
    public String toString() {
        return "MrpRezObj{" +
                "KosTr= '" + KosTr + '\'' +
                ", OfenTyp= '" + OfenTyp + '\'' +
                ", Zeile= " + Zeile +
                ", MatNr= '" + MatNr + '\'' +
                ", MatGrp= '" + MatGrp + '\'' +
                ", MatName= '" + MatName + '\'' +
                ", Rezeptanteil= " + Rezeptanteil +
                ", Preis= " + Preis +
                ", Ausbeute= " + Ausbeute +
                ", Monat= " + Monat +
                ", GattGrp= '" + GattGrp + '\'' +
                '}';
    }
}
