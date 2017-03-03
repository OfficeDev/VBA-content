---
title: DocumentLibraryVersion.Comments Property (Office)
keywords: vbaof11.chm277021
f1_keywords:
- vbaof11.chm277021
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.DocumentLibraryVersion.Comments
ms.assetid: ce99f474-527a-4895-c360-7e5d02435655
---


# DocumentLibraryVersion.Comments Property (Office)

Gets any optional comments associated with the specified version of the shared document. Read-only.


## Syntax

 _expression_. **Comments**

 _expression_ A variable that represents a **DocumentLibraryVersion** object.


## Remarks

A user can attach version comments through the document library user interface when checking in a document that was previously checked out.


## Example

The following example lists comments and other properties for each version of a shared document.


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

