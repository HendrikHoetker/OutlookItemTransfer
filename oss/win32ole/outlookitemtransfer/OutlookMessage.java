package oss.win32ole.outlookitemtransfer;

public class OutlookMessage {
	
	// holds the name of the file dropped from outlook
	private final String filename;
	
	
	// holds the binary data which then can be saved to a file on disk
	private final byte[] fileContents;

	
	public OutlookMessage(String filename, byte[] fileContents) {
		this.filename = filename;
		this.fileContents = fileContents;
	}


	public String getFilename() {
		return filename;
	}


	public byte[] getFileContents() {
		return fileContents;
	}

}
