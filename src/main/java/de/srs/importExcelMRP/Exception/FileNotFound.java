package de.srs.importExcelMRP.Exception;

public class FileNotFound extends Exception{

	/**
	 * 
	 */
	private static final long serialVersionUID = 8776370653129808621L;
	private String msg;
	public String getMsg() {
		return msg;
	}
	public void setMsg(String msg) {
		this.msg = msg;
	}
	public FileNotFound() {
		super();
	}
	public FileNotFound(String msg) {
		super();
		this.msg = msg;
	}
	
}
