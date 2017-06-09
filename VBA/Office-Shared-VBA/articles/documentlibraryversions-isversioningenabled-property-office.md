---
title: DocumentLibraryVersions.IsVersioningEnabled Property (Office)
keywords: vbaof11.chm277030
f1_keywords:
- vbaof11.chm277030
ms.prod: office
api_name:
- Office.DocumentLibraryVersions.IsVersioningEnabled
ms.assetid: 8f3035d5-9720-f87c-3b74-ef37f61b28bc
ms.date: 06/08/2017
---


# DocumentLibraryVersions.IsVersioningEnabled Property (Office)

Gets a  **Boolean** value that indicates whether the document library in which the active document is saved on the server is configured to create a backup copy, or version, each time the file is edited on the Web site. Read-only.


## Syntax

 _expression_. **IsVersioningEnabled**

 _expression_ A variable that represents a **DocumentLibraryVersions** object.


## Remarks

Versioning is enabled or disabled on the document library and not on individual documents. Therefore the value of the  **IsVersioningEnabled** property depends on the document library in which the document is saved.


## Example

The following example displays the number of saved versions of the active document, if versioning is enabled.


```
 Dim dlvVersions As Office.DocumentLibraryVersions 
 Set dlvVersions = ActiveDocument.DocumentLibraryVersions 
 If dlvVersions.IsVersioningEnabled Then 
 MsgBox "This document has " &amp; dlvVersions.Count &amp; _ 
 " saved versions.", vbInformation + vbOKOnly, _ 
 "Version Information" 
 Else 
 MsgBox "Versioning is not enabled for this document.", _ 
 vbInformation + vbOKOnly, "No Versioning" 
 End If 
 Set dlvVersions = Nothing 

```


## See also


#### Concepts


[DocumentLibraryVersions Object](documentlibraryversions-object-office.md)
#### Other resources


[DocumentLibraryVersions Object Members](documentlibraryversions-members-office.md)

