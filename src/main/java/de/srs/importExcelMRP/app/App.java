package de.srs.importExcelMRP.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.attribute.FileTime;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Timestamp;
import java.util.HashMap;
import java.util.Objects;
import java.util.Properties;
import java.util.TreeMap;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dhatim.fastexcel.reader.ExcelReaderException;

import de.srs.importExcelMRP.Exception.DatabaseException;
import de.srs.importExcelMRP.Exception.ErrorReadExcelException;
import de.srs.importExcelMRP.enums.ExcelEnum;
import de.srs.importExcelMRP.obj.MrpObj;
import de.srs.importExcelMRP.obj.MrpRezObjExt;

/**
 * TG Import_Excel_MRP
 *
 */
public class App {

    //wenn E ist wird die Zeile gelesen
    public static final String E = "E";
    protected static Connection instanceConn;

	protected static final  Logger logger = LogManager.getLogger(App.class);

	static String folderExcelFiles = "";
	//final static String fileName = "MRP25-TgAl-a-Arbdatei.xlsx";
	static String fileName = "";
	final static String SHEET_NAME = "Rezepte";
	final static int rowStart = 1;
	static int rowNumber = 0;

	private static final String END_OF_GIESS = "Einsatz GO o. Kr.Pprod";

	private static final String EXCEL_EXCEPTION = null;

    static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	static String projectpath;

	private static int rowCountSheet;

	private static MrpObj mrpObj;

	static HashMap<Integer, MrpRezObjExt> giessMonat1;
	static HashMap<Integer, MrpRezObjExt> giessMonat2;
	static HashMap<Integer, MrpRezObjExt> imcoMonat1;
	static HashMap<Integer, MrpRezObjExt> imcoMonat2;

	private static int emptyRow;

	private static boolean isNewKtrNummer;
	

	public static void main(String[] args) throws Exception {

			readFileProperties();
			readLastFileFronDisk();
			initialize();

			readKtr();

			try {
				instanceConn.close();
			} catch (SQLException e) {
                logger.error("Error:: {}", String.valueOf(e));
			}
			

	}

	public static void readFileProperties() throws IOException {
		

		String rootPath;
        rootPath = Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResource("")).getPath();
        String appConfigPath = rootPath + "app.properties";
        logger.debug("appConfigPath:: {}", appConfigPath);

		Properties appProps = new Properties();
		appProps.load(new FileInputStream(appConfigPath));


		String pathExcelFiles = appProps.getProperty("pathExcelFiles");
		folderExcelFiles = pathExcelFiles;
        logger.debug("pathExcelFiles:: {}", pathExcelFiles);
	}

	public static void readLastFileFronDisk() {
		TreeMap<Long, String> files = new TreeMap<>();

		final File folder = new File(folderExcelFiles);
	    for (final File fileEntry : Objects.requireNonNull(folder.listFiles())) {
	        if (fileEntry.isDirectory()) {
	        	readLastFileFronDisk();
	        } else {
	        	try {
	        		
	        	    FileTime creationTime = (FileTime) Files.getAttribute(fileEntry.getAbsoluteFile().toPath(), "creationTime");
	        	    //logger.info("File::"+ fileEntry.getName() + " "+ creationTime.toString() + " Ms::" + creationTime.toMillis());
	        	    files.put(creationTime.toMillis(), fileEntry.getName());
	        	    
	        	} catch (IOException ex) {
	        	    // handle exception
	        	}
	        }
	    }
        logger.info("The last file is:: {}", files.get(files.lastKey()));
	    //logger.info(files.keySet().toString());
	    fileName = files.get(files.lastKey());
	}

    /**
     * Initializes the application by reading an Excel file and setting related properties.
     * This method performs the following steps:
     * Combines folder and file name to construct the complete Excel file path.
     * Opens the specified Excel file using Apache POI and reads the configured sheet.
     * If an error occurs while opening the file or reading from it, an error message is logged.
     * Exceptions:
     * - IOException: If an error occurs while accessing the Excel file.
     */
	private static void initialize() {
		logger.info("Test Read File Excel");

		projectpath = System.getProperty("user.dir");
		logger.debug(projectpath);
        logger.debug("{}{}{}{}", folderExcelFiles, File.separator, File.separator, fileName);

		try {
			workbook = new XSSFWorkbook(folderExcelFiles + File.separator + File.separator + fileName);
			sheet = workbook.getSheet(SHEET_NAME);
			rowCountSheet = sheet.getPhysicalNumberOfRows();
		} catch (IOException e) {
            logger.error("This is a severe error message:: {}", String.valueOf(e));
		}
	}

	protected static void readKtr() throws ErrorReadExcelException {
		rowNumber = rowStart;
		mrpObj = new MrpObj();
		emptyRow = 0;

		deleteMrpTableContent();
		deleteMrpRezTableContent();
        deleteMatnrPreise();

		for (rowNumber ++; (rowNumber < rowCountSheet && emptyRow <4); rowNumber++) {
            logger.info("rowNumber:: {} emptyRow:: {}", rowNumber, emptyRow);
			isNewKtrNummer=false;

			try {
				if ( (sheet.getRow(rowNumber).getCell(ExcelEnum.KTRNR.getNr()) != null && sheet.getRow(rowNumber).getCell(ExcelEnum.KTRNR.getNr()).getCellType() != CellType.BLANK  &&
                        sheet.getRow(rowNumber).getCell(ExcelEnum.KTRBEZEICHUNG.getNr()) != null &&
                        sheet.getRow(rowNumber).getCell(ExcelEnum.KTRGRP.getNr()) != null && sheet.getRow(rowNumber).getCell(ExcelEnum.FORMAT.getNr()) != null) && emptyRow  < 4 ) {
					mrpObj.setKosTr(sheet.getRow(rowNumber).getCell(ExcelEnum.KTRNR.getNr()).getCellType().equals(CellType.NUMERIC) ||
                            sheet.getRow(rowNumber).getCell(ExcelEnum.KTRNR.getNr()).getCellType().equals(CellType.FORMULA) ? String.valueOf((int) sheet.getRow(rowNumber).getCell(ExcelEnum.KTRNR.getNr()).getNumericCellValue()) : sheet.getRow(rowNumber).getCell(ExcelEnum.KTRNR.getNr()).getStringCellValue());
					mrpObj.setKosTr_Bez(sheet.getRow(rowNumber).getCell(ExcelEnum.KTRBEZEICHUNG.getNr()).getStringCellValue().length() > 50
							? sheet.getRow(rowNumber).getCell(ExcelEnum.KTRBEZEICHUNG.getNr()).getStringCellValue().substring(0, 49)
							: sheet.getRow(rowNumber).getCell(1).getStringCellValue());
					mrpObj.setKostr_Grp(sheet.getRow(rowNumber).getCell(ExcelEnum.KTRGRP.getNr()).getStringCellValue());
					mrpObj.setFrmt(sheet.getRow(rowNumber).getCell(ExcelEnum.FORMAT.getNr()).getStringCellValue() != null
							? sheet.getRow(rowNumber).getCell(ExcelEnum.FORMAT.getNr()).getStringCellValue()
							: "");
					
					//mrpObj.setDatum_Zeit(Timestamp.from(Instant.now()));
					mrpObj.setDatum_Zeit(new Timestamp(System.currentTimeMillis()));

                    logger.info("RowNumber:: {} KTR - Bezeichung:: {}", rowNumber + 1, mrpObj.getKosTr_Bez());

					readRezAnteil(mrpObj);
					save();

				} else
					emptyRow++;

			} catch (Exception e) {
                logger.error("readKtr error:: {}", String.valueOf(e));
				throw new ErrorReadExcelException(EXCEL_EXCEPTION,"Failed to read excel row number:: " + rowNumber , e);
			}

		}

		logger.info(mrpObj.toString());
		logger.info("end");

	}



	protected static void readRezAnteil(MrpObj mrpObj) throws ErrorReadExcelException {
		// HashMap<HashMap<OfenEnum,MonatEnum>, HashMap<Integer, MrpRezObjExt>>
		// mrpRezDates = new HashMap<HashMap<OfenEnum,MonatEnum>,
		// HashMap<Integer,MrpRezObjExt>>();
		// HashMap<OfenEnum, MonatEnum> ofenMonat = new HashMap<OfenEnum,MonatEnum>();
		// HashMap<Integer, MrpRezObjExt> mrpRezDate = new
		// HashMap<Integer,MrpRezObjExt>();
		// ofenMonat.put(OfenEnum.GIESS, MonatEnum.MONAT1);

        emptyRow= 0;
		imcoMonat1 = new HashMap<>();
		imcoMonat2 = new HashMap<>();
		giessMonat1 = new HashMap<>();
		giessMonat2 = new HashMap<>();
		try {
			while (!isNewKtrNummer && emptyRow < 3) {
				rowNumber ++;
                logger.info("rowNumber:: {} EinsatzMat:: {}", rowNumber, sheet.getRow(rowNumber).getCell(ExcelEnum.EINSATZMATERIAL.getNr()));

                //toDo bessere kontrolle
                //Mat-nr E,Art.Art F ,EinsatzMaterial G
				if (getMatNr().isEmpty()) {
					emptyRow++;
				}
				else {
                    //toDo bessere kontrolle
					if ( (sheet.getRow(rowNumber).getCell(ExcelEnum.ZEILEEP.getNr()) != null) && (sheet.getRow(rowNumber).getCell(ExcelEnum.ZEILEEP.getNr()).getStringCellValue().equalsIgnoreCase(E))) {
						//Giess oder Imco Ofen ?
						if (sheet.getRow(rowNumber).getCell(ExcelEnum.EINSATZMATERIAL.getNr()).getStringCellValue().equals(END_OF_GIESS) || (sheet.getRow(rowNumber).getCell(ExcelEnum.EINSATZMATERIAL.getNr()) != null
                                && !String.valueOf(sheet.getRow(rowNumber).getCell(ExcelEnum.EINSATZMATERIAL.getNr()).getStringCellValue()).isEmpty() &&
                                sheet.getRow(rowNumber).getCell(ExcelEnum.ME_S_NOT_USED.getNr()) != null && sheet.getRow(rowNumber).getCell(ExcelEnum.ME_S_NOT_USED.getNr()).getStringCellValue().equalsIgnoreCase("t") ) ) {

                            imcoMonat1.put(imcoMonat1.size()+1, readRezAnteilImcoMonat1(mrpObj));
							imcoMonat2.put(imcoMonat2.size()+1, readRezAnteilImcoMonat2(mrpObj));
						}
						else {
                            giessMonat1.put(giessMonat1.size()+1, readRezAnteilGiessMonat1(mrpObj));
							giessMonat2.put(giessMonat2.size()+1, readRezAnteilGiessMonat2(mrpObj));
						}
					}
				}
			}
			emptyRow=0;
			
		} catch (Exception e) {
            logger.error("Error line nr.:: {} -- {}", rowNumber + 1, e);
			throw new ErrorReadExcelException(EXCEL_EXCEPTION,"Failed to read excel row number:: " + rowNumber , e);
			
		}

	}

    private static String getMatNr() {
        String matNr = "";
        if (sheet.getRow(rowNumber).getCell(ExcelEnum.MATNR.getNr()) != null) {
            if (sheet.getRow(rowNumber).getCell(4).getCellType().equals(CellType.NUMERIC) || sheet.getRow(rowNumber).getCell(4).getCellType().equals(CellType.FORMULA))
                matNr =String.valueOf((int) sheet.getRow(rowNumber).getCell(ExcelEnum.MATNR.getNr()).getNumericCellValue());
            else
                matNr = sheet.getRow(rowNumber).getCell(ExcelEnum.MATNR.getNr() ).getStringCellValue();
        }
        return matNr;
    }

    /**
	 * check if the line of excel is a part of Rezepte or the row is a new one: <br>
	 * ToDo Bitte auch KtrNr und Group und Format
	 * @return
	 * false if <b>KTR-Bezeichung</b> from Row +1 != null and currently <b>Mat.Gruppe</b> and <b>Einsatzmaterial</b> are not empty
	 */
	protected static boolean newKtrNummer() {

        return (sheet.getRow(rowNumber + 1).getCell(ExcelEnum.KTRBEZEICHUNG.getNr()) != null && sheet.getRow(rowNumber + 1).getCell(ExcelEnum.KTRBEZEICHUNG.getNr()).getCellType() != CellType.BLANK) &&
((sheet.getRow(rowNumber).getCell(5) != null && sheet.getRow(rowNumber).getCell(ExcelEnum.ARTART.getNr()).getStringCellValue().isEmpty()) &&
sheet.getRow(rowNumber).getCell(7) != null && sheet.getRow(rowNumber).getCell(ExcelEnum.EINSATZMATERIAL.getNr()).getStringCellValue().isEmpty());
		
	}

	/**
	 * Gieß Ofen aktueller Monat [1]
	 *
     *
	 */
	protected static MrpRezObjExt readRezAnteilGiessMonat1(MrpObj mrpObj) throws ErrorReadExcelException {
		MrpRezObjExt rpRezObjExt = new MrpRezObjExt();

        try {
			// _nr
			rpRezObjExt.setKosTr(mrpObj.getKosTr());
			// Mat-nr "E"
			if (sheet.getRow(rowNumber).getCell(4).getCellType().equals(CellType.NUMERIC) || sheet.getRow(rowNumber).getCell(ExcelEnum.MATNR.getNr()).getCellType().equals(CellType.FORMULA))
				rpRezObjExt.setMatNr(String.valueOf((int) sheet.getRow(rowNumber).getCell(ExcelEnum.MATNR.getNr()).getNumericCellValue()));
			else
				rpRezObjExt.setMatNr(sheet.getRow(rowNumber).getCell(ExcelEnum.MATNR.getNr()).getStringCellValue());

			// New Name Art.Art - Old name Mat-Grp "F" = 6
			rpRezObjExt.setMatGrp(sheet.getRow(rowNumber).getCell(ExcelEnum.ARTART.getNr()).getStringCellValue());
			// EMat "H" = 8
			rpRezObjExt.setEinsatzMaterial(sheet.getRow(rowNumber).getCell(ExcelEnum.EINSATZMATERIAL.getNr()).getStringCellValue());
			// AB "I" = 8
			rpRezObjExt.setAusbeute((float) sheet.getRow(rowNumber).getCell(ExcelEnum.AB.getNr()).getNumericCellValue());
			// Rezant-1 "J" = 9
			rpRezObjExt.setRezeptanteil((float) sheet.getRow(rowNumber).getCell(ExcelEnum.REZEPT1.getNr()).getNumericCellValue());
			// MI-1 "L" = 11
			rpRezObjExt.setMi(sheet.getRow(rowNumber).getCell(11) != null ? (float) sheet.getRow(rowNumber).getCell(ExcelEnum.MI1.getNr()).getNumericCellValue() : 0);
			// Moto-1 "N" = 13
			rpRezObjExt.setMoto1_2(sheet.getRow(rowNumber).getCell(ExcelEnum.Moto1.getNr()) != null
					? (float) sheet.getRow(rowNumber).getCell(ExcelEnum.Moto1.getNr()).getNumericCellValue()
					: 0);
			// IP-1 GO "P" = 15
			rpRezObjExt.setIp( sheet.getRow(rowNumber).getCell(ExcelEnum.IPGIESSOFEN1.getNr()) != null ? (float) sheet.getRow(rowNumber).getCell(ExcelEnum.IPGIESSOFEN1.getNr()).getNumericCellValue() : 0);
			// Preis-1 "T" = 19
			rpRezObjExt.setPreis((float) sheet.getRow(rowNumber).getCell(ExcelEnum.PREIS1.getNr()).getNumericCellValue());
			// MatKosten_rezeptindividuell-1 "V" = 21
			rpRezObjExt.setMatKosten(sheet.getRow(rowNumber).getCell(ExcelEnum.METKOSTEN1.getNr()) != null ? (float) sheet.getRow(rowNumber).getCell(ExcelEnum.METKOSTEN1.getNr()).getNumericCellValue() :0);
            //OfenTyp
            rpRezObjExt.setOfenTyp(sheet.getRow(rowNumber).getCell(ExcelEnum.OFENTYP.getNr()).getStringCellValue());
            //Gatt. Grp.
            rpRezObjExt.setGattGrp(sheet.getRow(rowNumber).getCell(ExcelEnum.ARTGATTGRP.getNr()).getStringCellValue());

            logger.info("RowNumber:: {} - Monat 1 -{}", rowNumber + 1, rpRezObjExt.toString());
			
			
		} catch (Exception e) {
			throw new ErrorReadExcelException(EXCEL_EXCEPTION,"readRezAnteilGiessMonat1 - Failed to read excel row number:: " + (rowNumber+1) , e);
		}

		return rpRezObjExt;
	}

	/**
	 * Gieß Monat 2
     */
	protected static MrpRezObjExt readRezAnteilGiessMonat2(MrpObj mrpObj) throws ErrorReadExcelException {

		MrpRezObjExt rpRezObjExt = new MrpRezObjExt();

		try {
			// _nr
			rpRezObjExt.setKosTr(mrpObj.getKosTr());
			// Mat-nr "E" = 5
			if (sheet.getRow(rowNumber).getCell(4).getCellType().equals(CellType.NUMERIC) || sheet.getRow(rowNumber).getCell(4).getCellType().equals(CellType.FORMULA))
				rpRezObjExt.setMatNr(String.valueOf((int) sheet.getRow(rowNumber).getCell(ExcelEnum.MATNR.getNr()).getNumericCellValue()));
			else
				rpRezObjExt.setMatNr(sheet.getRow(rowNumber).getCell(ExcelEnum.MATNR.getNr()).getStringCellValue());
			// Mat-Grp "F" = 6 --> Art.Art
			rpRezObjExt.setMatGrp(sheet.getRow(rowNumber).getCell(ExcelEnum.ARTART.getNr()).getStringCellValue());
			// EMat "G" = 7
			rpRezObjExt.setEinsatzMaterial(sheet.getRow(rowNumber).getCell(ExcelEnum.EINSATZMATERIAL.getNr()).getStringCellValue());
			// AB "I" = 8
			rpRezObjExt.setAusbeute((float) sheet.getRow(rowNumber).getCell(ExcelEnum.AB.getNr()).getNumericCellValue());
			// Rezant-2 "AC" = 29
			rpRezObjExt.setRezeptanteil((float) sheet.getRow(rowNumber).getCell(ExcelEnum.REZEPT2.getNr()).getNumericCellValue());
			// MI-2 "AE" = 31
			rpRezObjExt.setMi((float) sheet.getRow(rowNumber).getCell(ExcelEnum.MI2.getNr()).getNumericCellValue());
			// Moto-2 "AG" = 33
			rpRezObjExt.setMoto1_2(sheet.getRow(rowNumber).getCell(ExcelEnum.MOTO2.getNr()) != null
					? (float) sheet.getRow(rowNumber).getCell(ExcelEnum.MOTO2.getNr()).getNumericCellValue()
					: 0);
			// IP-2 GO "AI" = 35
			rpRezObjExt.setIp(sheet.getRow(rowNumber).getCell(ExcelEnum.IPGIESSOFEN2.getNr()) != null ? (float) sheet.getRow(rowNumber).getCell(ExcelEnum.IPGIESSOFEN2.getNr()).getNumericCellValue() :0);
			// Preis-2 "AM" = 39
			rpRezObjExt.setPreis((float) sheet.getRow(rowNumber).getCell(ExcelEnum.PREIS2.getNr()).getNumericCellValue());
			// MatKosten_rezeptindividuell-2 "A" = 41
			rpRezObjExt.setMatKosten(sheet.getRow(rowNumber).getCell(ExcelEnum.METKOSTEN2.getNr()) != null ? (float) sheet.getRow(rowNumber).getCell(ExcelEnum.METKOSTEN2.getNr()).getNumericCellValue() : 0);
            //OfenTyp
            rpRezObjExt.setOfenTyp(sheet.getRow(rowNumber).getCell(ExcelEnum.OFENTYP.getNr()).getStringCellValue());
            //Gatt. Grp.
            rpRezObjExt.setGattGrp(sheet.getRow(rowNumber).getCell(ExcelEnum.ARTGATTGRP.getNr()).getStringCellValue());

            logger.info("RowNumber:: {} - Monat 2 -{}", rowNumber + 1, rpRezObjExt.toString());
			
		} catch (Exception e) {
			throw new ErrorReadExcelException(EXCEL_EXCEPTION,"readRezAnteilGiessMonat2 - Failed to read excel row number:: " + rowNumber , e);
		}

		return rpRezObjExt;
	}

	/**
	 * Imco Ofen aktueller Monat [1]
	 *
     *
	 */
	protected static MrpRezObjExt readRezAnteilImcoMonat1(MrpObj mrpObj) throws ErrorReadExcelException {
		MrpRezObjExt rpRezObjExt = new MrpRezObjExt();
		try {
			//Art.Art Excel Datei Spalte F (6) und Einsatzmaterial G (7) not empty
			if (sheet.getRow(rowNumber).getCell(5) != null && !sheet.getRow(rowNumber).getCell(5).getStringCellValue().isEmpty() && sheet.getRow(rowNumber).getCell(7) != null && !sheet.getRow(rowNumber).getCell(7).getStringCellValue().isEmpty()) {
				// _nr
                setMrpObjCommon(mrpObj, rpRezObjExt);
                // ME-1 "J" = 10
				rpRezObjExt.setRezeptanteil((float) sheet.getRow(rowNumber).getCell(ExcelEnum.REZEPT1.getNr()).getNumericCellValue());
		        // MI-1 "K" = 11
				rpRezObjExt.setMi((float) sheet.getRow(rowNumber).getCell(ExcelEnum.MI1.getNr()).getNumericCellValue());
		        // Moto-1 "M" = 13
				rpRezObjExt.setMoto1_2((float) sheet.getRow(rowNumber).getCell(ExcelEnum.Moto1.getNr()).getNumericCellValue());
		        // IP GO-1 "R" = 17
				rpRezObjExt.setIp((float) sheet.getRow(rowNumber).getCell(ExcelEnum.IPIO1.getNr()).getNumericCellValue());
		        // Preis-1 "S" = 19
				rpRezObjExt.setPreis((float) sheet.getRow(rowNumber).getCell(ExcelEnum.PREIS1.getNr()).getNumericCellValue());
		        // MatKosten_rezeptindividuell-1 "W" = 22
				rpRezObjExt.setMatKosten(sheet.getRow(rowNumber).getCell(ExcelEnum.METKOSTEN1.getNr()) != null ? (float) sheet.getRow(rowNumber).getCell(ExcelEnum.METKOSTEN1.getNr()).getNumericCellValue() : 0);
				
			}
            logger.info("RowNumber:: {} - Imco  Monat 1 -{}", rowNumber + 1, rpRezObjExt.toString());
			
			
		} catch (Exception e) {
            logger.error("Error :: {}", String.valueOf(e));
			throw new ErrorReadExcelException(EXCEL_EXCEPTION," readRezAnteilImcoMonat1 - Failed to read excel row number:: " + rowNumber , e);
		}

		return rpRezObjExt;
	}

    /**
     * Sets common properties of the given MrpRezObjExt object based on values retrieved
     * from the specified MrpObj and other data sources.
     *
     * @param mrpObj The object containing source data for setting properties.
     * @param rpRezObjExt The object where the properties will be set.
     */
    private static void setMrpObjCommon(MrpObj mrpObj, MrpRezObjExt rpRezObjExt) {
        rpRezObjExt.setKosTr(mrpObj.getKosTr());
        // Mat-nr =  5
        rpRezObjExt.setMatNr(String.valueOf((int) sheet.getRow(rowNumber).getCell(ExcelEnum.MATNR.getNr()).getNumericCellValue()));
        // Art-Art "F" = 6
        rpRezObjExt.setMatGrp(sheet.getRow(rowNumber).getCell(ExcelEnum.ARTGATTGRP.getNr()).getStringCellValue());
        // EMat "G" = 7
        rpRezObjExt.setEinsatzMaterial(sheet.getRow(rowNumber).getCell(ExcelEnum.EINSATZMATERIAL.getNr()).getStringCellValue());
        // AB "I" = 8
        rpRezObjExt.setAusbeute((float) sheet.getRow(rowNumber).getCell(ExcelEnum.AB.getNr()).getNumericCellValue());
        //OfenTyp
        rpRezObjExt.setOfenTyp(sheet.getRow(rowNumber).getCell(ExcelEnum.OFENTYP.getNr()).getStringCellValue());
        //Gatt. Grp.
        rpRezObjExt.setGattGrp(sheet.getRow(rowNumber).getCell(ExcelEnum.ARTGATTGRP.getNr()).getStringCellValue());
    }

    /**
	 * Imco Ofen aktueller Monat [2]
	 *
     *
	 */
	protected static MrpRezObjExt readRezAnteilImcoMonat2(MrpObj mrpObj) throws ErrorReadExcelException {
		MrpRezObjExt rpRezObjExt = new MrpRezObjExt();
		try {
			//Art.Art Excel Datei Spalte F (6) und Einsatzmaterial H (7) not empty
			if (sheet.getRow(rowNumber).getCell(ExcelEnum.ARTART.getNr()) != null && !sheet.getRow(rowNumber).getCell(ExcelEnum.ARTART.getNr()).getStringCellValue().isEmpty() &&
                    sheet.getRow(rowNumber).getCell(ExcelEnum.EINSATZMATERIAL.getNr()) != null && !sheet.getRow(rowNumber).getCell(ExcelEnum.EINSATZMATERIAL.getNr()).getStringCellValue().isEmpty()) {
				// _nr
                setMrpObjCommon(mrpObj, rpRezObjExt);
                // ME-2 "AD" = 29
				rpRezObjExt.setRezeptanteil((float) sheet.getRow(rowNumber).getCell(ExcelEnum.ME_AE.getNr()).getNumericCellValue());

		        // MI-2 "AE" = 30
				rpRezObjExt.setMi((float) sheet.getRow(rowNumber).getCell(ExcelEnum.MI2.getNr()).getNumericCellValue());

		        // Moto-2 "AH" = 32
				rpRezObjExt.setMoto1_2((float) sheet.getRow(rowNumber).getCell(ExcelEnum.MOTO2.getNr()).getNumericCellValue());

		        // IP GO-2 "AM" = 36
				rpRezObjExt.setIp(sheet.getRow(rowNumber).getCell(ExcelEnum.IPIO2.getNr()) != null ? (float) sheet.getRow(rowNumber).getCell(ExcelEnum.IPIO2.getNr()).getNumericCellValue() : 0);

		        // Preis-2 "AN" = 38
				rpRezObjExt.setPreis((float) sheet.getRow(rowNumber).getCell(ExcelEnum.PREIS2.getNr()).getNumericCellValue());

		        // MatKosten_rezeptindividuell-1 "AP" = 40
				rpRezObjExt.setMatKosten(sheet.getRow(rowNumber).getCell(ExcelEnum.METKOSTEN2.getNr()) != null ?  (float) sheet.getRow(rowNumber).getCell(ExcelEnum.METKOSTEN2.getNr()).getNumericCellValue() : 0);
			}

            logger.info("RowNumber:: {} - Imco  Monat 2 -{}", rowNumber + 1, rpRezObjExt.toString());
			
		} catch (Exception e) {
			throw new ErrorReadExcelException(EXCEL_EXCEPTION," readRezAnteilImcoMonat2 - Failed to read excel row number:: " + rowNumber , e);
		}

		return rpRezObjExt;
	}

    protected static Connection getSqlDbConnection() {
		/*
		 * String jdbcUrl =
		 * "jdbc:sqlserver://DE-TG-SQL01\\SQLEXPRESS;databaseName=Giesserei"; String
		 * username = "real_sa"; String password = "sqlserver_sa";
		 */

		String connectionUrl = "jdbc:sqlserver://DE-TG-SQL01\\SQLEXPRESS:51767;" + "database=Giesserei;"
				+ "user=real_sa;" + "password=sqlserver_sa;" + "encrypt=true;" + "trustServerCertificate=true;"
				+ "loginTimeout=30;";

		if (instanceConn == null) {
			try {
				// Connection conn = DriverManager.getConnection(jdbcUrl, username, password)
                return (DriverManager.getConnection(connectionUrl));
			} catch (SQLException e) {
				logger.error("Error in getSqlDbConnection {} {} ", e.getSQLState(), e.getMessage());
			}

		}
		return instanceConn;

	}

	private static void save() throws DatabaseException {
		try {
			saveMrpInDB();
			saveMrpRezepteGiess();
			saveMrpRezepteImco();
            saveMatNrPreise();
		} catch (Exception e) {
			throw new DatabaseException("Save Error:: "+ e, e.getMessage());
		}
	}

    private static void saveMatNrPreise() throws DatabaseException {
        // Tabelle Matnr-Preise füllen
        // sql_str := 'insert into MRP_Mat_Preise (Matnr, Preis) select distinct(Matnr), Preis from MRP_Rezept';
        // 27.03.2025 neues Feld Monat in Tabelle MRP_Mat_Preise ergänzt
        instanceConn = getSqlDbConnection();

        String insertSql = "insert into MRP_Mat_Preise_TEST (Matnr, Preis, Monat) select distinct(Matnr), Preis, Monat from MRP_Rezept_TEST";
        PreparedStatement pstmt;
        try {
            pstmt = instanceConn.prepareStatement(insertSql);
            pstmt.execute();

        } catch (SQLException e) {
            logger.error("Error when I try to saveMatNrPreise. SQL State: {} {}", e.getSQLState(), e);
            throw new DatabaseException(e.getSQLState() , "Row Nr.: " + (rowNumber+1) + " - KTR Bez.: '" + mrpObj.getKosTr_Bez()+   "' Error when I try to saveMatNrPreise " + e);
        }

    }

    protected static void saveMrpInDB() throws DatabaseException {

		instanceConn = getSqlDbConnection();
		String insertSql = "insert into MRP_TEST (KosTr,KosTr_Bez,KosTr_Grp,Frmt,Datum_Zeit) "
				+ " VALUES( ?, ?, ?, ?, ?)";

		PreparedStatement pstmt;
		try {
			pstmt = instanceConn.prepareStatement(insertSql);
			pstmt.setString(1, mrpObj.getKosTr());
			pstmt.setString(2, mrpObj.getKosTr_Bez());
			pstmt.setString(3, mrpObj.getKostr_Grp().length() > 5 ? mrpObj.getKostr_Grp().substring(0,5) : mrpObj.getKostr_Grp());
			pstmt.setString(4, mrpObj.getFrmt());
			//TEST
			//pstmt.setString(5, mrpObj.getDatum_Zeit().toString().replace("-", ""));
			pstmt.setTimestamp(5, mrpObj.getDatum_Zeit());
			

			pstmt.execute();

		} catch (SQLException e) {
			logger.error("Error when I try to saveMrpInDB. SQL State: {} {}", e.getSQLState(), e);
			throw new DatabaseException(e.getSQLState() , "Row Nr.: " + (rowNumber+1) + " - KTR Bez.: '" + mrpObj.getKosTr_Bez()+   "' Error when I try to saveMrpInDB " + e);
		}

	}

    private static void saveMrpRezepte() throws DatabaseException {
        instanceConn = getSqlDbConnection();
        String insertSql = "insert into MRP_Rezept_TEST (KosTr,OfenTyp,Zeile,MatNr,ArtikelArt,MatName,Rezeptanteil,Preis,Ausbeute,Monat, MopsMatGrp)" +
                "Values (?,?,?,?,?,?,?,?,?,?,?)";

        PreparedStatement pstmt;
        try {
            pstmt = instanceConn.prepareStatement(insertSql);
            int s1 = giessMonat1.size();
            for (int i = 1; i <= s1; i++) {
                pstmt.setString(1, giessMonat1.get(i).getKosTr());
                pstmt.setString(2, giessMonat1.get(i).getOfenTyp());
                pstmt.setInt(3, i);
                pstmt.setString(4, giessMonat1.get(i).getMatNr());
                pstmt.setString(5, giessMonat1.get(i).getMatGrp());
                // ToDO delete
                //pstmt.setString(6, giessMonat1.get(i).getMatName());
                pstmt.setString(6, giessMonat1.get(i).getEinsatzMaterial());
                pstmt.setDouble(7, giessMonat1.get(i).getRezeptanteil()*100);
                pstmt.setDouble(8, giessMonat1.get(i).getPreis());
                pstmt.setDouble(9, giessMonat1.get(i).getAusbeute()*100);
                pstmt.setDouble(10, 1);
                pstmt.setString(11, giessMonat1.get(i).getGattGrp());


                pstmt.execute();

                pstmt.setString(1, giessMonat2.get(i).getKosTr());
                pstmt.setString(2, giessMonat1.get(i).getOfenTyp());
                pstmt.setInt(3, i);
                pstmt.setString(4, giessMonat2.get(i).getMatNr());
                pstmt.setString(5, giessMonat2.get(i).getMatGrp());
                // ToDO delete
                //pstmt.setString(6, giessMonat1.get(i).getMatName());
                pstmt.setString(6, giessMonat2.get(i).getEinsatzMaterial());
                pstmt.setDouble(7, giessMonat2.get(i).getRezeptanteil()*100);
                pstmt.setDouble(8, giessMonat2.get(i).getPreis());
                pstmt.setDouble(9, giessMonat2.get(i).getAusbeute()*100);
                pstmt.setDouble(10, 2);
                pstmt.setString(11, giessMonat2.get(i).getGattGrp());

                pstmt.execute();

            }

        }
        catch (SQLException e) {
            throw new DatabaseException(e.getSQLState() , "Row Nr.: " + (rowNumber+1) + " - KTR Bez.: '" + mrpObj.getKosTr_Bez()+   "' Error when I try to saveMrpRezepteGiess " + e);
        }
        catch (Exception e) {
            throw new ExcelReaderException(" Error when I try to saveMrpRezepteGiess:: " + e);
        }
    }
	private static void saveMrpRezepteGiess() throws DatabaseException {
		instanceConn = getSqlDbConnection();
		String insertSql = "insert into MRP_Rezept_TEST (KosTr,OfenTyp,Zeile,MatNr,ArtikelArt,MatName,Rezeptanteil,Preis,Ausbeute,Monat, MopsMatGrp)" +
                         "Values (?,?,?,?,?,?,?,?,?,?,?)";

        String t = "";

		PreparedStatement pstmt;
		try {
			pstmt = instanceConn.prepareStatement(insertSql);
			 int s1 = giessMonat1.size();
			 for (int i = 1; i <= s1; i++) {
                 pstmt.setString(1, giessMonat1.get(i).getKosTr());
				 pstmt.setString(2, giessMonat1.get(i).getOfenTyp());
				 pstmt.setInt(3, i);
				 pstmt.setString(4, giessMonat1.get(i).getMatNr());
				 pstmt.setString(5, giessMonat1.get(i).getMatGrp());
				 // ToDO delete
				 //pstmt.setString(6, giessMonat1.get(i).getMatName());
				 pstmt.setString(6, giessMonat1.get(i).getEinsatzMaterial());
				 pstmt.setDouble(7, giessMonat1.get(i).getRezeptanteil()*100);
				 pstmt.setDouble(8, giessMonat1.get(i).getPreis());
				 pstmt.setDouble(9, giessMonat1.get(i).getAusbeute()*100);
				 pstmt.setDouble(10, 1);
                 pstmt.setString(11, giessMonat1.get(i).getGattGrp());


				 pstmt.execute();

				 pstmt.setString(1, giessMonat2.get(i).getKosTr());
				 pstmt.setString(2, giessMonat1.get(i).getOfenTyp());
				 pstmt.setInt(3, i);
				 pstmt.setString(4, giessMonat2.get(i).getMatNr());
				 pstmt.setString(5, giessMonat2.get(i).getMatGrp());
				 // ToDO delete
				 //pstmt.setString(6, giessMonat1.get(i).getMatName());
				 pstmt.setString(6, giessMonat2.get(i).getEinsatzMaterial());
				 pstmt.setDouble(7, giessMonat2.get(i).getRezeptanteil()*100);
				 pstmt.setDouble(8, giessMonat2.get(i).getPreis());
				 pstmt.setDouble(9, giessMonat2.get(i).getAusbeute()*100);
				 pstmt.setDouble(10, 2);
                 pstmt.setString(11, giessMonat2.get(i).getGattGrp());

				 pstmt.execute();

			}

		}
		catch (SQLException e) {
			throw new DatabaseException(e.getSQLState() , "Row Nr.: " + (rowNumber+1) + " - KTR Bez.: '" + mrpObj.getKosTr_Bez()+   "' Error when I try to saveMrpRezepteGiess " + e);
		}
		catch (Exception e) {
			throw new ExcelReaderException(" Error when I try to saveMrpRezepteGiess:: " + e);
		}
	}

	private static void saveMrpRezepteImco() throws DatabaseException {
		instanceConn = getSqlDbConnection();
		String insertSql = "insert into MRP_Rezept_TEST (KosTr,OfenTyp,Zeile,MatNr,ArtikelArt,MatName,Rezeptanteil,Preis,Ausbeute,Monat,MopsMatGrp)" +
                         "Values (?,?,?,?,?,?,?,?,?,?,?)";

		PreparedStatement pstmt;
		try {
			pstmt = instanceConn.prepareStatement(insertSql);
			 int s1 = imcoMonat1.size();
			 for (int i = 1; i <= s1; i++) {

                 pstmt.setString(1, imcoMonat1.get(i).getKosTr());
				 pstmt.setString(2, imcoMonat1.get(i).getOfenTyp());
				 pstmt.setInt(3, i);
				 pstmt.setString(4, imcoMonat1.get(i).getMatNr());
				 pstmt.setString(5, imcoMonat1.get(i).getMatGrp());
				 // ToDO delete
				 //pstmt.setString(6, imcoMonat1.get(i).getMatName());
				 pstmt.setString(6, imcoMonat1.get(i).getEinsatzMaterial());
				 pstmt.setDouble(7, imcoMonat1.get(i).getRezeptanteil()*100);
				 pstmt.setDouble(8, imcoMonat1.get(i).getPreis());
				 pstmt.setDouble(9, imcoMonat1.get(i).getAusbeute()*100);
				 //ToDo delete Monat
				 pstmt.setDouble(10, 1);
                 pstmt.setString(11, imcoMonat1.get(i).getGattGrp());

				 pstmt.execute();
				 pstmt.setString(1, imcoMonat2.get(i).getKosTr());
				 //ToDo delete
				 pstmt.setString(2, imcoMonat2.get(i).getOfenTyp());
				 //pstmt.setString(2, "I");
				 pstmt.setInt(3, i);
				 pstmt.setString(4, imcoMonat2.get(i).getMatNr());
				 pstmt.setString(5, imcoMonat2.get(i).getMatGrp());
				 // ToDO delete
				 //pstmt.setString(6, giessMonat1.get(i).getMatName());
				 pstmt.setString(6, imcoMonat2.get(i).getEinsatzMaterial());
				 pstmt.setDouble(7, imcoMonat2.get(i).getRezeptanteil()*100);
				 pstmt.setDouble(8, imcoMonat2.get(i).getPreis());
				 pstmt.setDouble(9, imcoMonat2.get(i).getAusbeute()*100);
				 //ToDo delete
				 pstmt.setDouble(10, 2);
                 pstmt.setString(11, imcoMonat2.get(i).getGattGrp());

				 pstmt.execute();
			}

		}
		catch (SQLException e) {
			throw new DatabaseException(e.getSQLState() , "Row Nr.: " + (rowNumber+1) + " - KTR Bez.: '" + mrpObj.getKosTr_Bez()+   "' Error when I try to saveMrpRezepteImco " + e);
		}
		catch (Exception e) {
			throw new ExcelReaderException("Error when I try to saveMrpRezepteImco:: " + e);
		}
		
	}	
	protected static void deleteMrpTableContent() {
		instanceConn = getSqlDbConnection();
		String sqlString = "DELETE FROM MRP_TEST";

        deleteDatesFromTables(sqlString);
    }
	protected static void deleteMrpRezTableContent() {
		instanceConn = getSqlDbConnection();
		String sqlString = "DELETE FROM MRP_Rezept_TEST";

        deleteDatesFromTables(sqlString);
    }
    protected static void deleteMatnrPreise() {
        instanceConn = getSqlDbConnection();
        String sqlString = "DELETE FROM MRP_Mat_Preise_TEST";

        deleteDatesFromTables(sqlString);
    }

    protected static void deleteDatesFromTables(String sqlString) {
        try {
            Statement stmt = instanceConn.createStatement();
            stmt.executeUpdate(sqlString);
        } catch (SQLException e) {
            try {
                instanceConn.close();
            } catch (SQLException e1) {
                logger.error("Erro when I try to close db connection");
            }

            logger.error("Error when I try to deleteMrpTableContent. SQL State: {} {}", e.getSQLState(),       e.getMessage());
        }
    }


}
