package oss.win32ole.outlookitemtransfer;

import org.eclipse.swt.dnd.Transfer;
import org.eclipse.swt.dnd.TransferData;

public abstract class OutlookItemTransfer extends Transfer {
	
	abstract public boolean isSupportedType(TransferData transferData);
	abstract public Object nativeToJava(TransferData transferData);

}
