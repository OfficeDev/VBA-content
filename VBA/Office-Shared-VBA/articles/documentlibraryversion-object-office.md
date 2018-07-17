---
title: DocumentLibraryVersion Object (Office)
keywords: vbaof11.chm277014
f1_keywords:
- vbaof11.chm277014
ms.prod: office
api_name:
- Office.DocumentLibraryVersion
ms.assetid: ac13975d-4f91-1fc5-5b0a-94b21309ffb7
ms.date: 06/08/2017
---


# DocumentLibraryVersion Object (Office)

The  **DocumentLibraryVersion** object represents a single saved version of a shared document which has versioning enabled and which is stored in a document library on the server. Each **DocumentLibraryVersion** object is a member of the active document's **DocumentLibraryVersions** collection.


## Remarks

 Each **DocumentLibraryVersion** object represents one saved version of the active document. When versioning is enabled, a new version is created on the server when the actions listed below occur; additional versions are not created each time the user saves changes to the open document.


- Check In
    
- Save - A new version is created on the server when the user first saves the document after opening it. Additional changes saved while the document is open apply to the same version.
    
- Restore
    
- Upload
    


Use the  **Modified**, **ModifiedBy**, and **Comments** properties to return information about a saved version of a shared document.

Use the  **Open** method to open a previous version, or the **Restore** method to restore a previous version in place of the current version. Use the **Delete** method to delete a version.


## Example

The following example displays the properties of each saved version of the active document.


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


## Methods



|**Name**|
|:-----|
|[Delete](documentlibraryversion-delete-method-office.md)|
|[Open](documentlibraryversion-open-method-office.md)|
|[Restore](documentlibraryversion-restore-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](documentlibraryversion-application-property-office.md)|
|[Comments](documentlibraryversion-comments-property-office.md)|
|[Creator](documentlibraryversion-creator-property-office.md)|
|[Index](documentlibraryversion-index-property-office.md)|
|[Modified](documentlibraryversion-modified-property-office.md)|
|[ModifiedBy](documentlibraryversion-modifiedby-property-office.md)|
|[Parent](documentlibraryversion-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
