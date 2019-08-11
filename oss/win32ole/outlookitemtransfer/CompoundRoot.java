package oss.win32ole.outlookitemtransfer;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

class CompoundRoot extends CompoundContainer {
	
	/**
	 * constructor
	 */
	CompoundRoot() {
		super("root", OutlookStorageType.Root);
	}


	/**
	 * exports the compound structure to a byte array using APACHE POI library
	 * this is doing all the magic to dump the dragged data from outlook into a bytestream
	 * the stream can be written to a filestream and used as a usual .msg file.
	 */
	byte[] toByteArray() {
		
		try {
			// create an empty File System
			POIFSFileSystem fs = new POIFSFileSystem();
			
			// append all sub elements to file system root
			writeDirectory(this, fs.getRoot());
			
			// used to extract the binary data from the file system object
			ByteArrayOutputStream bos = new ByteArrayOutputStream();
			
			// write file system to the byte array stream
			// a fileoutputstream could be used, too. Check the documentation of POIFSFileSystem
			fs.writeFilesystem(bos);
			
			// close the file system object
			fs.close();
			
			
			// close the byte array output stream and extract the byte data
			bos.flush();
			byte[] result = bos.toByteArray();
			bos.close();
			
			return result;
		} catch (Exception e) {}
		
		// if something went wrong, return null
		return null;
	}


	private void writeDirectory(CompoundObject storage, DirectoryEntry dir) {
		try {
			switch (storage.getType()) {
				case Root:
					for (CompoundObject subElement: ((CompoundRoot)storage).getSubElements()) {
						writeDirectory(subElement, dir);
					}
					break;
					
				case Storage:
					DirectoryEntry subDir = dir.createDirectory(storage.getName());
					for (CompoundObject subElement: ((CompoundStorage)storage).getSubElements()) {
						writeDirectory(subElement, subDir);
					}
					break;
					
				case Stream:
					dir.createDocument(storage.getName(), new ByteArrayInputStream(((CompoundStream)storage).getData()));
					break;
					
				default:
					break;
			}
		} catch (Exception e) {}
	}

	
	
	public String toString() {
		return toStringInternal("   ");
	}
	
	
	protected String toStringInternal(String indent) {
		String result = String.format("Root <%s>, Count Sub-Elements: %d\n", getName(), getSubElements().size());
		
		for (CompoundObject sub: getSubElements()) {
			result += sub.toStringInternal(indent + "   ");
		}
		
		return result + "\n";
	}

}
