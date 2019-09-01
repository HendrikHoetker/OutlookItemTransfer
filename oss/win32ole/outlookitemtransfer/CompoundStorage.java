package oss.win32ole.outlookitemtransfer;

import org.apache.commons.codec.binary.StringUtils;
import org.eclipse.swt.internal.ole.win32.COM;
import org.eclipse.swt.internal.ole.win32.IEnumSTATSTG;
import org.eclipse.swt.internal.ole.win32.IStorage;
import org.eclipse.swt.internal.ole.win32.STATSTG;
import org.eclipse.swt.internal.win32.OS;

class CompoundStorage extends CompoundContainer {
	
	/**
	 * constructor
	 */
	public CompoundStorage(String name) {
		super(name, OutlookStorageType.Storage);
	}
	
	
	/**
	 * creates from pointer to an IStorage object
	 */
	public static CompoundRoot createFromStorage(long pIStorage) {
		if (pIStorage == 0) {
			return null;
		}
		
	 	CompoundRoot root = new CompoundRoot();
		readOutlookStorage(root, pIStorage);
		
		return root;
	}
	
	
	/**
	 * reads the contents of the IStorage and appends accordingly to the parent object
	 * this method is calling itself recursively to walk through the whole IStorage object
	 * read contents of the IStorage which can be either IStorages or IStreams
	 * works same as directory / file in a filesystem
	 */
	private static void readOutlookStorage(CompoundContainer parent, long pIStorage) {
		// open IStorage object
		IStorage storage = new IStorage(pIStorage);
		storage.AddRef();
		
		
		// walk through the content of the IStorage object
		long[] pEnumStorage = new long[1];
		if (storage.EnumElements(0, 0, 0, pEnumStorage) == COM.S_OK) {
			
			// get storage iterator
			IEnumSTATSTG enumStorage = new IEnumSTATSTG(pEnumStorage[0]);
			enumStorage.AddRef();
			enumStorage.Reset();
			
			// prepare statstg structure which tells about the object found by the iterator
			long pSTATSTG = OS.GlobalAlloc(OS.GMEM_FIXED | OS.GMEM_ZEROINIT, STATSTG.sizeof);
			int[] fetched = new int[1];
			
			while (enumStorage.Next(1, pSTATSTG, fetched) == COM.S_OK && fetched[0] == 1) {
				
				// get the description of the the object found
				STATSTG statstg = new STATSTG();
				COM.MoveMemory(statstg, pSTATSTG, STATSTG.sizeof);
				
				// get the name of the object found
				String name = readPWCSName(statstg);

				// depending on type of object
				switch (statstg.type) {
					case COM.STGTY_STREAM: {	// load an IStream (=File)
						try {
							long[] pIStream = new long[1];
							
							// get the pointer to the IStream
							if (storage.OpenStream(name, 0, COM.STGM_DIRECT | COM.STGM_READ | COM.STGM_SHARE_EXCLUSIVE, 0, pIStream) == COM.S_OK) {
								// load the IStream
								CompoundStream stream = CompoundStream.createFromStream(name, pIStream[0]);
								if (stream != null) {
									parent.addSubElement(stream);
								}
							}
						} catch (Exception e) {}
					}
					break;
					
					case COM.STGTY_STORAGE: {	// load an IStorage (=SubDirectory)
						try {
							long[] pSubIStorage = new long[1];
							
							// get the pointer to the sub IStorage
							if (storage.OpenStorage(name, 0, COM.STGM_DIRECT | COM.STGM_READ | COM.STGM_SHARE_EXCLUSIVE, null, 0, pSubIStorage) == COM.S_OK) {
								
								// recursively walk through the sub storage
								CompoundStorage subStorage = new CompoundStorage(name);
								parent.addSubElement(subStorage);
								
								// recursive call to this function for sub storage
								readOutlookStorage(subStorage, pSubIStorage[0]);
							}
						} catch (Exception e) {}
					}
					break;
				}
			}
			
			// close the iterator
			enumStorage.Release();
		}
		
		// close the IStorage object
		storage.Release();
	}
	
	
	/**
	 * reads the name of the object in the IStorage structure
	 * = name of directory of name of file
	 */
	private static String readPWCSName(STATSTG statstg) {
		byte[] name = new byte[64];
		COM.MoveMemory(name, statstg.pwcsName, name.length);
		
		// note that microsoft uses wide char names (=UTF-16 Little Endian)
		String result = StringUtils.newStringUtf16Le(name);
		
		// the wide char has a \0 to indicate string end, StringUtils does not recognize it
		// so the string is truncated after the first \0
		result = result.substring(0, result.indexOf(new String(new byte[] {0})));
		
		return result;
	}



	
	protected String toStringInternal(String indent) {
		String result = String.format("\n%sStorage <%s>, Count Sub-Elements: %d\n", indent, getName(), getSubElements().size());
		
		for (CompoundObject sub: getSubElements()) {
			result += sub.toStringInternal(indent + "   ");
		}
		
		return result + "\n";
	}
	
}
