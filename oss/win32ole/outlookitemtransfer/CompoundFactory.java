package oss.win32ole.outlookitemtransfer;

class CompoundFactory {
	
	
	/**
	 * creates a root object containing one stream
	 */
	static CompoundRoot createFromIStream(long pIStream) {
		CompoundRoot root = new CompoundRoot();
		
		CompoundStream stream = CompoundStream.createFromStream("unknown", pIStream);
		if (stream != null) {
			root.addSubElement(stream);
		}
		
		return root;
	}
	
	
	/**
	 * creates a root object containing all sub  elements of the IStorage
	 */
	static CompoundRoot createFromIStorage(long pIStorage) {
		return CompoundStorage.createFromStorage(pIStorage);
	}


}
