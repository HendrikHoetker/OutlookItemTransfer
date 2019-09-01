package oss.win32ole.outlookitemtransfer;

import com.bosch.pmtoolbox.utils.OSType;

public class OutlookItemTransferFactory {
	
	public static OutlookItemTransfer getOutlookItemTransfer() {
		if (OSType.get() == OSType.Windows) {
			return OutlookItemTransferImpl.getInstance();
		} 

		return null;
	}

}
