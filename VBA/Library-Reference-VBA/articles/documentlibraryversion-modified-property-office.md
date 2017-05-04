---
title: DocumentLibraryVersion.Modified Property (Office)
keywords: vbaof11.chm277017
f1_keywords:
- vbaof11.chm277017
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.DocumentLibraryVersion.Modified
ms.assetid: 3bcf1913-cdc9-36b9-7548-9804b56411e1
---


# DocumentLibraryVersion.Modified Property (Office)

Gets the date and time at which the specified version of the shared document was last saved to the server. Read-only.


## Syntax

 _expression_. **Modified**

 _expression_ A variable that represents a **DocumentLibraryVersion** object.


## Remarks

A new version is created on the server each time a user opens the document and is updated when the user saves changes; additional versions are not created each time the user saves changes to the open document. The  **Modified** property of the active document version represents the last time the user saved changes to the open document.


## Example

The following example displays the Modified date and time along with other properties of each version of a shared document.


```vb
 Dim dlvVersions As Office.DocumentLibraryVersions 
 Dim dlvVersion As Office.DocumentLibraryVersion 
 Dim strVersionInfo As String 
 Set dlvVersions = ActiveDocument.DocumentLibraryVersions 
 If dlvVersions.IsVersioningEnabled Then 
 strVersionInfo = "This document has " &; _ 
 dlvVersions.Count &; " versions: " &; vbCrLf 
 For Each dlvVersion In dlvVersions 
 strVersionInfo = strVersionInfo &; _ 
 " - Version #: " &; dlvVersion.Index &; vbCrLf &; _ 
 " - Modified by: " &; dlvVersion.ModifiedBy &; vbCrLf &; _ 
 " - Modified on: " &; dlvVersion.Modified &; vbCrLf &; _ 
 " - Comments: " &; dlvVersion.Comments &; vbCrLf 
 Next 
 Else 
 strVersionInfo = "Versioning not enabled for this document." 
 End If 
 MsgBox strVersionInfo, vbInformation + vbOKOnly, "Version Information" 
 Set dlvVersion = Nothing 
 Set dlvVersions = Nothing
```


## See also


#### Concepts


[DocumentLibraryVersion Object](documentlibraryversion-object-office.md)

