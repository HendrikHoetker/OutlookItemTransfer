package oss.win32ole.outlookitemtransfer;

import org.eclipse.swt.dnd.TransferData;

public class OutlookItemTransferStub extends OutlookItemTransfer {
	
	private OutlookItemTransferStub() {}

	private static OutlookItemTransferStub instance = null;
	public static OutlookItemTransferStub getInstance() {
		if (instance == null) {
			instance = new OutlookItemTransferStub();
		}
		
		return instance;
	}
	
	/**
	 * only supports dragged IStorage objects
	 */
	@Override
	public TransferData[] getSupportedTypes() {
		return null;
	}
	
	/**
	 * this will listen on two drag events:
	 * - filegroupdescriptor
	 * - file contents
	 */
	@Override
	protected int[] getTypeIds() {
		return null;
	}
	
	
	/**
	 * get names of drag events
	 */
	@Override
	protected String[] getTypeNames() {
		return null;
	}
	

	/**
	 * check if this transfer supports the data
	 */
	@Override
	public boolean isSupportedType(TransferData transferData) {
		return false;
	}
	
	
	/**
	 * not implemented as drag to outlook is not considered
	 */
	@Override
	public void javaToNative(Object arg0, TransferData arg1) {
	}
	
	
	/**
	 * converts dragged emails from outlook to a list of outlookmessages
	 */
	@Override
	public Object nativeToJava(TransferData transferData) {
		return null;
	}

}
