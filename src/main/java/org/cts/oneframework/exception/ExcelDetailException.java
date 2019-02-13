package org.cts.oneframework.exception;

public class ExcelDetailException extends RuntimeException {

	private static final long serialVersionUID = 1L;

	public ExcelDetailException() {
		super();
	}

	public ExcelDetailException(String arg0, Throwable arg1, boolean arg2, boolean arg3) {
		super(arg0, arg1, arg2, arg3);
	}

	public ExcelDetailException(String arg0, Throwable arg1) {
		super(arg0, arg1);
	}

	public ExcelDetailException(String arg0) {
		super(arg0);
	}

	public ExcelDetailException(Throwable arg0) {
		super(arg0);
	}
}
