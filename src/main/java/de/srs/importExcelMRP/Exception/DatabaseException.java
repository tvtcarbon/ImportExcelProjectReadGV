package de.srs.importExcelMRP.Exception;

public class DatabaseException extends Exception {
    /**
	 * 
	 */
	private static final long serialVersionUID = -7045830044972946702L;
	private String code;

    public DatabaseException(String code, String message) {
        super(message);
        this.setCode(code);
    }

    public DatabaseException(String code, String message, Throwable cause) {
        super(message, cause);
        this.setCode(code);
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

}
