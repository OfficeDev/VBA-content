---
title: DocumentLibraryVersion.Modified Property (Office)
keywords: vbaof11.chm277017
f1_keywords:
- vbaof11.chm277017
ms.prod: office
api_name:
- Office.DocumentLibraryVersion.Modified
ms.assetid: 3bcf1913-cdc9-36b9-7548-9804b56411e1
ms.date: 06/08/2017
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


```
 Dim dlvVersions As Office.DocumentLibraryVersions 
 Dim dlvVersion As Office.DocumentLibraryVersion 
 Dim strVersionInfo As String 
 Set dlvVersions = ActiveDocument.DocumentLibraryVersions 
 If dlvVersions.IsVersioningEnabled Then 
 strVersionInfo = "This document has " &amp; _ 
 dlvVersions.Count &amp; " versions: " &amp; vbCrLf 
 For Each dlvVersion In dlvVersions 
 strVersionInfo = strVersionInfo &amp; _ 
 " - Version #: " &amp; dlvVersion.Index &amp; vbCrLf &amp; _ 
 " - Modified by: " &amp; dlvVersion.ModifiedBy &amp; vbCrLf &amp; _ 
 " - Modified on: " &amp; dlvVersion.Modified &amp; vbCrLf &amp; _ 
 " - Comments: " &amp; dlvVersion.Comments &amp; vbCrLf 
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
#### Other resources


[DocumentLibraryVersion Object Members](documentlibraryversion-members-office.md)

