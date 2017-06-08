---
title: Document.CheckIn Method (Word)
keywords: vbawd10.chm158007645
f1_keywords:
- vbawd10.chm158007645
ms.prod: word
api_name:
- Word.Document.CheckIn
ms.assetid: 3c0e5a48-65e1-c7f7-c9f1-cabaabdcb3cb
ms.date: 06/08/2017
---


# Document.CheckIn Method (Word)

Returns a document from a local computer to a server, and sets the local document to read-only so that it cannot be edited locally.


## Syntax

 _expression_ . **CheckIn**( **_SaveChanges_** , **_Comments_** , **_MakePublic_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SaveChanges_|Optional| **Boolean**| **True** saves the document to the server location. The default is **True** .|
| _Comments_|Optional| **Variant**|Comments for the revision of the document being checked in (only applies if SaveChanges equals  **True** ).|
|||||
| _MakePublic_|Optional| **Boolean**| **True** allows the user to perform a publish on the document after being checked in. This submits the document for the approval process, which can eventually result in a version of the document being published to users with read-only rights to the document (only applies if _SaveChanges_ equals **True** ). The default is **False** .|
|||||
|||||

## Remarks

To take advantage of the collaboration features built into Microsoft Word, documents must be stored on a Microsoft SharePoint Portal Server.


## Example

This example checks the server to see if the specified document can be checked in. If it can be, it saves and closes the document and checks it back into the server.


```vb
Sub CheckInOut(docCheckIn As String) 
 If Documents(docCheckIn).CanCheckin = True Then 
 Documents(docCheckIn).CheckIn 
 MsgBox docCheckIn &; " has been checked in." 
 Else 
 MsgBox "This file cannot be checked in " &; 
 "at this time. Please try again later." 
 End If 
End Sub
```

To call the CheckInOut subroutine, use the following subroutine and replace  _"http://servername/workspace/report.doc"_ with the file name of an actual file located on the server mentioned in the Remarks section above.




```vb
Sub CheckDocInOut() 
 Call CheckInOut (docCheckIn:="http://servername/workspace/report.doc") 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

