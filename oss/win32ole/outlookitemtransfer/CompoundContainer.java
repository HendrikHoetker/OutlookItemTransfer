package oss.win32ole.outlookitemtransfer;

import java.util.ArrayList;
import java.util.List;

abstract class CompoundContainer extends CompoundObject {
	
	/**
	 * contains all elements of the root tree
	 */
	private List<CompoundObject> subElements;

	
	public CompoundContainer(String name, OutlookStorageType type) {
		super(name, type);
		this.subElements = new ArrayList<>();
	}


	/**
	 * adds a sub element to the root
	 */
	void addSubElement(CompoundObject element) {
		this.subElements.add(element);
	}
	
	
	/**
	 * get list of sub elements
	 */
	List<CompoundObject> getSubElements() {
		return this.subElements;
	}


}
