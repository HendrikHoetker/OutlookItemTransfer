package oss.win32ole.outlookitemtransfer;

import java.io.ByteArrayOutputStream;

import org.apache.commons.codec.binary.Hex;
import org.eclipse.swt.internal.ole.win32.COM;
import org.eclipse.swt.internal.ole.win32.IStream;
import org.eclipse.swt.internal.win32.OS;

class CompoundStream extends CompoundObject {
	
	/**
	 * contains the binary data (content of the compound file - IStream)
	 */
	private final byte[] data;

	
	/**
	 * constructor
	 */
	private CompoundStream(String name, byte[] data) {
		super(name, OutlookStorageType.Stream);
		
		this.data = data;
	}
	
	
	/**
	 * creates a compound stream from the pointer to an IStream object
	 */
	public static CompoundStream createFromStream(String name, long pIStream) {
		byte[] data = readStream(pIStream);
		
		if (data != null) {
			return new CompoundStream(name, data);
		}
		
		return null;
	}
	
	
	/**
	 * reads the data from an IStream object
	 */
	private static byte[] readStream(long pIStream) {
		
		// verify if valid pointer is given
		if (pIStream == 0) {
			return null;
		}
		
		try {
			// create and open the IStream
			IStream stream =  new IStream(pIStream);
			stream.AddRef();
			
			// opens the byte array output writer to collect all byte array chunks from the stream
			ByteArrayOutputStream bos = new ByteArrayOutputStream();

			// read 16k per iteration
			final int chunkSize = 16384;
			
			// monitors how many bytes actually are written
			int[] bytesWritten = new int[1];
			
			// contains the pointer of each chunk
			long pv = COM.CoTaskMemAlloc(chunkSize);

			while (stream.Read(pv, chunkSize, bytesWritten) == COM.S_OK && bytesWritten[0] > 0) {
				// reserve memory for the chunk data
				byte[] buffer = new byte[bytesWritten[0]];
				
				// move the chunk data to the byte array
				OS.MoveMemory(buffer, pv, bytesWritten[0]);
				
				// append to the byte array writer
				bos.write(buffer);
			}
			
			// finished writing
			bos.flush();
			
			// copy result from the byte array writer
			byte[] result = bos.toByteArray();
			bos.close();
			
			// close IStream object
			stream.Release();
			
			return result;
		} catch (Exception  e) {
			
			// on any exception return null
			return null;
		}
	}
	
	
	/**
	 * get the binary data earlier read from the IStream
	 */
	public byte[] getData() {
		return this.data;
	}
	
	
	/**
	 * returns a string of the object
	 */
	protected String toStringInternal(String indent) {
		if (this.data != null && this.data.length > 0) {
			return String.format("%sStream <%s> Length: %8d\tData: %s\n", indent, getName(), getData().length, Hex.encodeHexString(this.data, false));
		} else {
			return String.format("%sStream <%s> Length: %8d\tData: %s\n", indent, getName(), 0, "<empty>");
		}
	}

}
