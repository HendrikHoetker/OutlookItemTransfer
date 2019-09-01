package oss.win32ole.outlookitemtransfer;

class CompoundFactory {
	
	
	/**
	 * returns a CompoundStream object based on an IStream
	 * @param pIStream
	 * @return
	 */
	static CompoundStream createFromIStream(long pIStream) {
		CompoundStream stream = CompoundStream.createFromStream("unknown", pIStream);
		if (stream != null) {
			return stream;
		}
		
		return null;
	}
	
	
	/**
	 * creates a root object containing all sub  elements of the IStorage
	 */
	static CompoundRoot createFromIStorage(long pIStorage) {
		return CompoundStorage.createFromStorage(pIStorage);
	}


}
