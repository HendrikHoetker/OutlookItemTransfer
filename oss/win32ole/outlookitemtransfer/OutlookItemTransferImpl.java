package oss.win32ole.outlookitemtransfer;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.codec.binary.StringUtils;
import org.eclipse.swt.dnd.TransferData;
import org.eclipse.swt.internal.ole.win32.COM;
import org.eclipse.swt.internal.ole.win32.FORMATETC;
import org.eclipse.swt.internal.ole.win32.IDataObject;
import org.eclipse.swt.internal.ole.win32.STGMEDIUM;
import org.eclipse.swt.internal.win32.OS;

public class OutlookItemTransferImpl extends OutlookItemTransfer {
	
	// additional constants needed, not provided by the COM interface
	private final int TYMED_ISTREAM = 4;
	private final int TYMED_ISTORAGE = 8;
	
	// the way outlook provides the dragged data
	private final int FILE_DESCRIPTOR_SIZE = 592;
	private final int FILE_DESCRIPTOR_START_INDEX = 72;
	
	
	/**
	 * singleton of the transfer
	 */
	private OutlookItemTransferImpl() {}
	private static OutlookItemTransferImpl instance = null;
	public static OutlookItemTransferImpl getInstance() {
		if (instance == null) {
			instance = new OutlookItemTransferImpl();
		}

		return instance;
	}
	
	
	/**
	 * only supports dragged IStorage objects
	 */
	@Override
	public TransferData[] getSupportedTypes() {
		int[] typeIds = getTypeIds();
		TransferData[] result = new TransferData[typeIds.length];
		
		for (int i = 0; i < typeIds.length; i++) {
			result[i] = new TransferData();
			result[i].type = typeIds[i];
			result[i].formatetc = new FORMATETC();
			result[i].formatetc.cfFormat = typeIds[i];
			result[i].formatetc.dwAspect = COM.DVASPECT_CONTENT;
			result[i].formatetc.lindex = -1;
			
			// it seems that older versions of outlook also provide direct IStreams or HGlobals
			// verifed on Outlook 2016, providing only IStorage objects
			result[i].formatetc.tymed = TYMED_ISTORAGE;
		}
		
		return result;
	}
	
	
	/**
	 * this will listen on two drag events:
	 * - filegroupdescriptor
	 * - file contents
	 */
	@Override
	protected int[] getTypeIds() {
		return new int[] {registerType(getTypeNames()[0]), registerType(getTypeNames()[1])};
	}
	
	
	/**
	 * get names of drag events
	 */
	@Override
	protected String[] getTypeNames() {
		// register for unicode UTF-16LE file name and file contents
		return new String[] {"FileGroupDescriptorW", "FileContents"};
	}
	
	
	/**
	 * check if this transfer supports the data
	 */
	@Override
	public boolean isSupportedType(TransferData transferData) {
		if (transferData == null) {
			return false;
		}

		for (int type: getTypeIds()) {
			final boolean checkTypeId = type == transferData.formatetc.cfFormat;
			final boolean checkAspect = (transferData.formatetc.dwAspect & COM.DVASPECT_CONTENT) == COM.DVASPECT_CONTENT;
			final boolean checkTymed = ((transferData.formatetc.tymed & COM.TYMED_HGLOBAL) == COM.TYMED_HGLOBAL) ||
					((transferData.formatetc.tymed & TYMED_ISTORAGE) == TYMED_ISTORAGE) ||
					((transferData.formatetc.tymed & TYMED_ISTREAM) == TYMED_ISTREAM);
			
			return checkTypeId && checkAspect && checkTymed;
		}
		
		return false;
	}
	
	
	/**
	 * not implemented as drag to outlook is not considered
	 */
	@Override
	protected void javaToNative(Object arg0, TransferData arg1) {
	}
	
	
	/**
	 * converts dragged emails from outlook to a list of outlookmessages
	 */
	@Override
	public Object nativeToJava(TransferData transferData) {
		if (transferData == null) {
			return null;
		}
		
		if (!isSupportedType(transferData)) {
			return null;
		}
		
		if (transferData.pIDataObject == 0) {
			return null;
		}
		
		
		// result list
		List<OutlookMessage> draggedMessages = new ArrayList<>();
		
		
		// open object of dragged data
		IDataObject dataObject = new IDataObject(transferData.pIDataObject);
		dataObject.AddRef();
		
		
		// contains raw data retrieved from IDataObject
		STGMEDIUM medium = new STGMEDIUM();
		medium.tymed = transferData.formatetc.tymed;
		if (dataObject.GetData(transferData.formatetc, medium) != COM.S_OK) {
			dataObject.Release();
			return null;
		}
		
		// get pointer to the file descriptor contained in medium.unionField
		long fileDescriptorPointer = OS.GlobalLock(medium.unionField);

		// get file count dragged from outlook
		int[] pFileCount = new int[1];
		OS.MoveMemory(pFileCount, fileDescriptorPointer, 4);
		
		// advance file descriptor pointer to the file descriptor structs
		fileDescriptorPointer += 4;
		
		// walk over the single file descriptors
		for (int i = 0; i < pFileCount[0]; i++) {
			// get the filename of the file
			String filename = getFilename(fileDescriptorPointer);

			// get the file contents of file index
			byte[] fileContents = getFileContents(i, dataObject);
			
			draggedMessages.add(new OutlookMessage(filename, fileContents));
			
			if (i < pFileCount[0] - 1) {
				// advance to the next file descriptor if this is not the last element
				// otherwise an invalid pointer exception is thrown not recognized by the java vm
				fileDescriptorPointer += FILE_DESCRIPTOR_SIZE;
			}
		}

		// release the medium
		OS.GlobalFree(medium.unionField);
		
		// close the IDataObject
		dataObject.Release();
		
		return draggedMessages.toArray();
	}
	
	
	/**
	 * get the name of the file from the file descriptor
	 * @param fileDescriptorPointer
	 * @return
	 */
	private String getFilename(long fileDescriptorPointer) {
		// the actual filename is located after an offset of 72 bytes
		byte[] filenameBytes = new byte[FILE_DESCRIPTOR_SIZE];
		OS.MoveMemory(filenameBytes, fileDescriptorPointer + FILE_DESCRIPTOR_START_INDEX, FILE_DESCRIPTOR_SIZE);
		
		// parse wide char bytes to java string
		String filename = StringUtils.newStringUtf16Le(filenameBytes);
		filename = filename.trim();
		return filename;
	}


	/**
	 * get the file contents of the given file
	 * parses the IStorage object down to an array of bytes
	 * representing the byte array of the file dragged
	 */
	private byte[] getFileContents(int fileIndex, IDataObject data) {
		FORMATETC format = new FORMATETC();
		format.cfFormat = getTypeIds()[1];
		format.dwAspect = COM.DVASPECT_CONTENT;
		format.lindex = fileIndex;
		format.ptd = 0;
		format.tymed = TYMED_ISTORAGE | TYMED_ISTREAM | COM.TYMED_HGLOBAL;
		
		STGMEDIUM medium = new STGMEDIUM();

		if (data.GetData(format, medium) == COM.S_OK) {
			switch (medium.tymed) {
				case TYMED_ISTORAGE: {
					CompoundRoot root = CompoundFactory.createFromIStorage(medium.unionField);
					if (root != null) {
						return root.toByteArray();
					}
				}

				default:
					// HGLOBAL or ISTREAM is not implemented
					return null;
			}
		}
		
		return null;
	}

}
