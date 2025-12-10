package de.srs.importExcelMRP.Exception;

public class ErrorReadExcelException extends Exception {
    /**
	 * 
	 */
	private static final long serialVersionUID = 2058363262098365491L;
	private String code;

    public ErrorReadExcelException(String code, String message) {
        super(message);
        this.setCode(code);
    }

    public ErrorReadExcelException(String code, String message, Throwable cause) {
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
