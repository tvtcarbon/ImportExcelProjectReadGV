package de.srs.importExcelMRP.enums;

public enum ExcelEnum {

    //*** Table MRP *******
    //Materialnumber
	KTRNR (0,"A", null, MonatEnum.MONAT1),
	//Änderung
	OFENTYP (1, "B", null, null),
    //Materialdescription
	KTRBEZEICHUNG (3, "D", null, MonatEnum.MONAT1),
	//Order
	KTRGRP (4, "E", null, MonatEnum.MONAT1),
	FORMAT(5, "F", null, MonatEnum.MONAT1),
    //*********************

    //***** Table MRP_Rezept *******************************
	MATNR(6, "G", null, MonatEnum.MONAT1),
	MATGRUPPE(7, "H", null, MonatEnum.MONAT1),
	ARTART(8, "I", null, MonatEnum.MONAT1),
	GATTGRP(9, "J", null, MonatEnum.MONAT1),
	EINSATZMATERIAL(12, "M", null, MonatEnum.MONAT1),
	AB(13, "N", null, MonatEnum.MONAT1),

    //************ RezeptAnteil ****************************************
	//************ Monat 1 *********************************************
	REZEPT_ANTEIL_GIESS_M1(14, "O", OfenEnum.GIESS, MonatEnum.MONAT1),
	REZEPT_ANTEIL_IMCO_M1(15,"P", OfenEnum.IMCO, MonatEnum.MONAT1),
	MI1(16, "Q", null, MonatEnum.MONAT1),
	T_1(17, "R", null, MonatEnum.MONAT1),
	Moto1(18, "S", null, MonatEnum.MONAT1),
	ME_T_NOT_USED(19, "T", null, null),
	IPGIESSOFEN1(20, "U", OfenEnum.GIESS, MonatEnum.MONAT1),
	ME_V_NOT_USED(21,"V", null, MonatEnum.MONAT1),
	IPIO1(22, "W", OfenEnum.IMCO, MonatEnum.MONAT1),
	ME_X_NOT_USED(23, "X", null, null),
	PREIS1(24, "Y", null, MonatEnum.MONAT1),
	ME_Z_NOT_USED (25, "Z", null, null),
	METKOSTEN1 (26, "AA", OfenEnum.GIESS, MonatEnum.MONAT1),
	IOKOSTEN1(27, "AB", OfenEnum.IMCO, MonatEnum.MONAT1),
	ME_AC_NOT_USED (28, "AC", null,null),
	ABGO1 (29, "AD", OfenEnum.GIESS, MonatEnum.MONAT1),
    ABIO(30, "AE", OfenEnum.IMCO, MonatEnum.MONAT1),

	//Einsatz
	EINSATZ(88, "CK", null,null),

	//************ Monat 2 *********************************************
	REZEPT_ANTEIL_GIESS_M2(31, "AF", OfenEnum.GIESS, MonatEnum.MONAT2),
	REZEPT_ANTEIL_IMCO_M2(32, "AG", OfenEnum.IMCO, MonatEnum.MONAT2),
	MI2(33, "AH", null, MonatEnum.MONAT2),
	MOTO2(35, "AJ",null, null),
	IPGIESSOFEN2(37, "AL", OfenEnum.GIESS, MonatEnum.MONAT2),
	IPIO2 (39, "AN", OfenEnum.IMCO, MonatEnum.MONAT2),
    PREIS2(41, "AP", null, MonatEnum.MONAT2 ),
    METKOSTEN2(43, "AR", OfenEnum.GIESS, MonatEnum.MONAT2),
    IOKOSTEN2(44, "AS", OfenEnum.IMCO, MonatEnum.MONAT2),
    ABGO2(46, "AU", OfenEnum.GIESS, MonatEnum.MONAT2),
    ABIO2(47, "AV", OfenEnum.IMCO, MonatEnum.MONAT2);


	private final int columnNumber;
	private final String columnName;
	private final OfenEnum ofenTyp;
	private final MonatEnum monat;
	
	ExcelEnum(int columnNumber, String columnName, OfenEnum ofenTyp, MonatEnum monat) {
		this.columnNumber = columnNumber;
		this.columnName = columnName;
		this.ofenTyp = ofenTyp;
		this.monat = monat;
	}

	public String getColumnName() {
		return columnName;
	}

	public int getNr() {
		return columnNumber;
	}

    public MonatEnum getMonat() {
		return monat;
	}


}
