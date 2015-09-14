/**
 * Owner: ShaownS
 * File: MessageException.java
 * Package: org.fusecensus.poireader
 * Project: FUSECensus
 * Email: ssarker@ncsu.edu
 */
package org.fusecensus.poireader;

/**
 * Simple message exception class to return my error messages
 * that makes a bit more sense than a stack trace.
 */
public class MessageException extends Exception {
	
	public MessageException(String message) {
		super(message);
		this.message = message;
	}
	
	public MessageException(Throwable cause) {
		super(cause);
	}
	
	public MessageException(String message, Throwable cause) {
		super(message, cause);
		this.message = message;
	}
	
	public MessageException(String message, Throwable cause, boolean enableSuppression, 
			boolean writeStacktrace) {
		super(message, cause, enableSuppression, writeStacktrace);
		this.message = message;
	}
	
	@Override
	public String toString() {
		return this.message;
	}
	
	@Override
	public String getMessage() {
		return this.message;
	}
	
	private static final long serialVersionUID = 1L;
	private String message = null;
}
