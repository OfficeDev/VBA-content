---
title: Document.CanCheckin Method (Word)
keywords: vbawd10.chm158007647
f1_keywords:
- vbawd10.chm158007647
ms.prod: word
api_name:
- Word.Document.CanCheckin
ms.assetid: 7021b14b-3e45-9850-bc59-d76c267f2934
ms.date: 06/08/2017
---


# Document.CanCheckin Method (Word)

 **True** if Microsoft Word can check in a specified document to a server. Read/write **Boolean** .


## Syntax

 _expression_ . **CanCheckin**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Return Value

Boolean


## Remarks

To take advantage of the collaboration features built into Word, documents must be stored on a Microsoft SharePoint Portal Server.


## Example

This example checks the server to see if the specified document can be checked in and, if it can be, closes the document and checks it back into the server.


```vb
Sub CheckInOut(docCheckIn As String) 
 If Documents(docCheckIn).CanCheckin = True Then 
 Documents(docCheckIn).CheckIn 
 MsgBox docCheckIn &; " has been checked in." 
 Else 
 MsgBox "This file cannot be checked in " &; _ 
 "at this time. Please try again later." 
 End If 
End Sub
```

To call the CheckInOut subroutine above, use the following subroutine and replace the "http://servername/workspace/report.doc" file name with an actual file located on a server mentioned in the Remarks section above.




```vb
Sub CheckDocInOut() 
 Call CheckInOut (docCheckIn:="http://servername/workspace/report.doc") 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

