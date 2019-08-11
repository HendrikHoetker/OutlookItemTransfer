# OutlookItemTransfer

## Motivation
There are some closed source libs on the market which are quite costly to implement Drag'n Drop from Outlook to a Java application.
The are also many complicated ways including using JNI or writing the dragged item to a temp file and read it back again.
I wanted a small library to be used same as other Transfers in Java SWT (like FileTransfer or LocalSelectionTransfer).
Target was not to provide the dragged outlook item as byte array.

## License
All source code files are provided under MIT License.

## Description
I provide all code required for implementing Drag'n Drop from Outlook to an Eclipse SWT application using the SWT drag'n drop processes.


## Dependencies to other libraries
The implementation relies on Apache Commons (StringUtils) and Apache POI Library. Both dependencies must be added to your project, too.


## Implementation

### Validation
The implementation is validated on
* Windows 10
* Outlook Professional 2016
* Eclipse 2019.3 Java EE

### OutlookItemTransfer
OutlookItemTransfer provides the Drag implementation (only dragging from Outlook, but not dropping to Outlook is implemented)
