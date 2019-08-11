# OutlookItemTransfer

## Motivation
There are some closed source libs on the market which are quite costly to implement Drag'n Drop from Outlook to a Java application. There are also many sources which are very incomplete or misleading in the internet.

Also a lot of statments in the internet (stackexchange.com, etc.) tell about very complicated ways including "doing it in C++ or C#", using JNI and/or writing the dragged item to a temp file and read it back again.

I wanted a small library to be used the same way as other Transfer classes in Java SWT (like FileTransfer or LocalSelectionTransfer). My target was to provide the dragged Outlook item as a an array of bytes.

## Description
I provide all code required for implementing Drag'n Drop from Outlook to an Eclipse SWT application using the SWT drag'n drop processes.

## Dependencies to other libraries
The implementation relies on:
* Apache Commons (StringUtils) for creating UTF16-LE strings, standard Java String could also be used but requires additional try...catch.
* Apache POI Library for writing the resulting MSG file to a byte stream.
Both dependencies must be added to your project, too.


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
	
The method javaToNative is not implemented as dropping to Outlook is not considered in my use cases.
  
### OutlookItem
The OutlookItem is representing one single item dragged from Outlook. The OutlookItem provides
* the filename to the file (as given by outlook)
* a byte[] to the contents of the file dragged from Outlook.
 
### Internals
The dragged item from Outlook is provided as IStorage object. IStorage is similar to a file system directory containing further IStorage objects (sub directories) as well as IStream objects representing files. 
 
The internal structure is called CompoundObject where an IStorage is represented by an CompoundStorage, an IStream by an CompoundStream. Hint: I added a CompoundRoot to indicate the file system's root directory. Outlook provides an IStorage as root object.
 
Outlook provides a virtual file system of IStorages and IStreams which are 1:1 represented in a tree of CompoundStorage and CompoundStreams.

The data of the IStream is extracted and stored at each CompoundStream object.

The .MSG file is nothing else than a binary dump of the virtual file system plus some header informations. This is the same also for other Office files (.XLS, .DOC, etc.).

The resulting virtual file system is actually written using the Apache POI library and the byte stream is extracted.

In the end the implementation could be simplified by removing the Compound* objects and directly writing the information to the Apache POI library. Intention was here to have a 1:1 representation of the dragged object in memory for further analyzes. One could use CompoundRoot.toString() to dump the initially provided IStorage data to a String.

## License
All source code files are provided under MIT License, see LICENSE.md.
I do not provide any warranty to the code. It is only tested under limited conditions for my own use cases.

## Resources
* https://stackoverflow.com/questions/7690236/can-i-drag-items-from-outlook-into-my-swt-application
* https://en.wikipedia.org/wiki/Component_Object_Model
* https://docs.microsoft.com/en-us/windows/win32/api/objidl/nn-objidl-istorage
* https://docs.microsoft.com/de-de/windows/win32/api/objidl/nn-objidl-istream
* https://poi.apache.org/apidocs/dev/org/apache/poi/poifs/filesystem/POIFSFileSystem.html
* http://www.fileformat.info/format/outlookmsg/
* https://www.eclipse.org/forums/index.php/t/164289/
* https://support.microsoft.com/de-de/help/171907/info-save-message-to-msg-compound-file
* https://github.com/Baltimore99/delphi-drag-drop/blob/master/Demos/Outlook/OutlookTarget.pas
* https://www.tenouk.com/visualcplusmfc/visualcplusmfc26.html
* https://de.switch-case.com/52136961
* https://www.delphipraxis.net/172132-outlook-attachments-einem-tstream.html
* https://github.com/eclipse/eclipse.platform.swt/blob/master/bundles/org.eclipse.swt/Eclipse%20SWT%20PI/win32/org/eclipse/swt/internal/ole/win32/IStorage.java
