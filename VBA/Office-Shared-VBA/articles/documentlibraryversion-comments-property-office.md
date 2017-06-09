---
title: DocumentLibraryVersion.Comments Property (Office)
keywords: vbaof11.chm277021
f1_keywords:
- vbaof11.chm277021
ms.prod: office
api_name:
- Office.DocumentLibraryVersion.Comments
ms.assetid: ce99f474-527a-4895-c360-7e5d02435655
ms.date: 06/08/2017
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

