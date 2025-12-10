package de.srs.importExcelMRP.enums;

public enum ExcelEnum {

	KTRNR (0,"A", null, MonatEnum.MONAT1),
	KTRBEZEICHUNG (1, "B", null, MonatEnum.MONAT1),
	KTRGRP (2, "C", null, MonatEnum.MONAT1),
	FORMAT(3, "D", null, MonatEnum.MONAT1),
	MATNR(4, "E", null, MonatEnum.MONAT1),
	ARTART(5, "F", null, MonatEnum.MONAT1),
	ARTGATTGRP(6, "G", null, MonatEnum.MONAT1),
	EINSATZMATERIAL(7, "H", null, MonatEnum.MONAT1),
	AB(8, "I", null, MonatEnum.MONAT1),
	REZEPT1(9, "J", OfenEnum.GIESS, MonatEnum.MONAT1),
	ME(10,"K", OfenEnum.IMCO, MonatEnum.MONAT1),
	MI1(11, "L", null, MonatEnum.MONAT1),
	T(12, "M", null, MonatEnum.MONAT1),
	Moto1(13, "N", null, MonatEnum.MONAT1),
	ME_O_NOT_USED(14, "O", null, null),
	IPGIESSOFEN1(15, "P", OfenEnum.GIESS, MonatEnum.MONAT1),
	ME_Q_NOT_USED(16,"Q", null, MonatEnum.MONAT1),
	IPIO1(17, "R", OfenEnum.IMCO, MonatEnum.MONAT1),
	ME_S_NOT_USED(18, "S", null, null),
	PREIS1(19, "T", null, MonatEnum.MONAT1),
	ME_U_NOT_USED (20, "U", null, null),
	METKOSTEN1 (21, "V", OfenEnum.GIESS, MonatEnum.MONAT1),
	IOKOSTEN1(22, "W", OfenEnum.IMCO, MonatEnum.MONAT1),
	ME_X_NOT_USED (23, "X", null,null),
	ABGO1 (24, "Y", OfenEnum.GIESS, MonatEnum.MONAT1),
    ABIO(25, "Z", OfenEnum.IMCO, MonatEnum.MONAT1),
	ZEILEEP (26, "AA", null,null),
	MATHAUPTGRP (27, "AB", null, null),
	OFENTYP (28, "AC", null, null),
	
	REZEPT2(29, "AD", OfenEnum.GIESS, MonatEnum.MONAT2),
	ME_AE(30, "AE", OfenEnum.IMCO, MonatEnum.MONAT2),
	MI2(31, "AF", null, MonatEnum.MONAT2),
	T_AG_NOT_USED(32, "AG", null, null),
	MOTO2(33, "AH",null, null),
	ME_AI_NOT_USED (34, "AI", null, null),
	IPGIESSOFEN2(35, "AJ", OfenEnum.GIESS, MonatEnum.MONAT2),
	ME_AK_NOT_USED (36, "AK", null, null),
	IPIO2 (37, "AL", OfenEnum.IMCO, MonatEnum.MONAT2),
    ME_AM_NOT_USED(38, "AM", null, null),
    PREIS2(39, "AN", null, MonatEnum.MONAT2 ),
    ME_AO_NOT_USED(40, "AO", null, null),
    METKOSTEN2(41, "AP", OfenEnum.GIESS, MonatEnum.MONAT2),
    IOKOSTEN2(42, "AQ", OfenEnum.IMCO, MonatEnum.MONAT2),
    ME_AR_NOT_USED(43, "AR", null, null),
    ABGO2(44, "AS", OfenEnum.GIESS, MonatEnum.MONAT2),
    ABIO2(45, "AT", OfenEnum.IMCO, MonatEnum.MONAT2);


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
