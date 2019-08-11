# OutlookItemTransfer

## Motivation
There are some closed source libs on the market which are quite costly to implement Drag'n Drop from Outlook to a Java application.
The are also many complicated ways including using JNI or writing the dragged item to a temp file and read it back again.
I wanted a small library to be used same as other Transfers in Java SWT (like FileTransfer or LocalSelectionTransfer).
Target was not to provide the dragged outlook item as byte array.

## License
All source code files are provided under MIT License, see LICENSE.md.
I do not provide any warranty to the code. It is only tested under limited conditions for my own use cases.

## Description
I provide all code required for implementing Drag'n Drop from Outlook to an Eclipse SWT application using the SWT drag'n drop processes.

## Dependencies to other libraries
The implementation relies on Apache Commons (StringUtils) and Apache POI Library. Both dependencies must be added to your project, too.


## Implementation

### Usage
			
	  if (OutlookItemTransfer.getInstance().isSupportedType(event.currentDataType)) {
		  Object o = OutlookItemTransfer.getInstance().nativeToJava(event.currentDataType);
		  if (o != null && o instanceof List) {
			  //...
		  }
	  }


### Validation
The implementation is validated on
* Windows 10
* Outlook Professional 2016
* Eclipse 2019.3 Java EE

### OutlookItemTransfer
OutlookItemTransfer provides the Drag implementation (only dragging from Outlook, but not dropping to Outlook is implemented). The Transfer class is subclassed to Transfer, not ByteArrayTransfer. The behaviour is very similar to the ByteArrayTransfer, but it considers some small modifications compared to the ByteArrayTransfer.

The OutlookItemTransfer can be used as any other SWT Transfer class.

The method nativeToJava will return a List of OutlookItem: List<OutlookItem>. Each item represents one dragged item from Outlook.
  
### OutlookItem
The OutlookItem is representing one single item dragged from Outlook. The OutlookItem provides
* the filename to the file (as given by outlook)
* a byte[] to the contents of the file dragged from Outlook.
 
### Internals
The dragged item from outlook is provided as IStorage object. IStorage is similar to a file system directory containing further IStorage objects (sub directories) as well as IStream objects representing files.
 
The internal structure is called CompoundObject where an IStorage is represented by an CompoundStorage, an IStream by an CompoundStream. Hint: I added a CompoundRoot to indicate the file system's root directory. Outlook provides an IStorage as root object.
 
Outlook provides a virtual file system of IStorages and IStreams which are 1:1 represented in a tree of CompoundStorage and CompoundStreams.

The data of the IStream is extracted and stored at each CompoundStream object.

The .MSG file is nothing else than a binary dump of the virtual file system plus some header informations. This is the same also for other Office files (.XLS, .DOC, etc.).

The resulting virtual file system is actually written using the Apache POI library and the byte stream is extracted.

In the end the implementation could be simplified by removing the Compound* objects and directly writing the information to the Apache POI library. Intention was here to have a 1:1 representation of the dragged object in memory for further analyzes. One could use CompoundRoot.toString() to dump the initially provided IStorage data to a String.

