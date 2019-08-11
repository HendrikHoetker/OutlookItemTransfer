package oss.win32ole.outlookitemtransfer;

abstract class CompoundObject {
	
	public static enum OutlookStorageType {
		Root,
		Stream,
		Storage
	}
	
	private final String name;
	private final OutlookStorageType type;
	
	
	CompoundObject(String name, OutlookStorageType type) {
		this.name = name;
		this.type = type;
	}


	protected String getName() {
		return this.name;
	}


	protected OutlookStorageType getType() {
		return this.type;
	}

	protected abstract String toStringInternal(String indent);
}
