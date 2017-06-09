---
title: DocumentLibraryVersions Object (Office)
keywords: vbaof11.chm277026
f1_keywords:
- vbaof11.chm277026
ms.prod: office
api_name:
- Office.DocumentLibraryVersions
ms.assetid: 075c0315-fade-6d45-9ab9-6c798f6f09ac
ms.date: 06/08/2017
---


# DocumentLibraryVersions Object (Office)

The  **DocumentLibraryVersions** property of the **Document** object in Microsoft Word, the **Workbook** object in Microsoft Excel, and the **Presentation** object in Microsoft PowerPoint returns a **DocumentLibraryVersions** object. The **DocumentLibraryVersions** object represents a collection of **DocumentLibraryVersion** objects.


## Remarks

Use the  **DocumentLibraryVersions** object with documents stored in a SharePoint document library on the server to determine whether versioning is enabled for the active document and, if versioning is enabled, to manage the document's collection of **DocumentLibraryVersion** objects.

 Each **DocumentLibraryVersion** object represents one saved version of the active document. When versioning is enabled, a new version is created on the server when the actions listed below occur; additional versions are not created each time the user saves changes to the open document.


- Check In
    
- Save - A new version is created on the server when the user first saves the document after opening it. Additional changes saved while the document is open apply to the same version.
    
- Restore
    
- Upload
    


The  **DocumentLibraryVersions** object model is available whether versioning is enabled or disabled on the active document. The **DocumentLibraryVersions** property of the **Document**, **Workbook** and **Presentation** objects does not return **Nothing** when the active document is not stored in a document library or versioning is not enabled. Use the **IsVersioningEnabled** property to determine whether the document library is configured to save a backup copy, or version, each time the document is edited on the Web site.


## Example

The following example checks to see whether versioning is enabled for the active document and, if so, displays information about each saved version.


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


## Properties



|**Name**|
|:-----|
|[Application](documentlibraryversions-application-property-office.md)|
|[Count](documentlibraryversions-count-property-office.md)|
|[Creator](documentlibraryversions-creator-property-office.md)|
|[IsVersioningEnabled](documentlibraryversions-isversioningenabled-property-office.md)|
|[Item](documentlibraryversions-item-property-office.md)|
|[Parent](documentlibraryversions-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
