---
title: SharedWorkspaceLink.Notes Property (Office)
keywords: vbaof11.chm270003
f1_keywords:
- vbaof11.chm270003
ms.prod: office
api_name:
- Office.SharedWorkspaceLink.Notes
ms.assetid: 5bb05b61-2746-f276-5159-ee8f28a30c66
ms.date: 06/08/2017
---


# SharedWorkspaceLink.Notes Property (Office)

Gets or sets the optional notes associated with a shared workspace link. Read/write.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Notes**

 _expression_ A variable that represents a **SharedWorkspaceLink** object.


## Example

The following example creates a list of all the shared workspace links that contain the word "building" in the Notes field.


```
Dim strBuildingLinks As String 
Dim swsLink As Office.SharedWorkspaceLink 
For Each swsLink In ActiveWorkbook.SharedWorkspace.Links 
   If InStr(swsLink.Notes, "building", vbTextCompare) > 0 Then 
      strBuildingLinks = strBuildingLinks &amp; swsLink.Description &amp; vbCrLf 
   End If 
Next 
MsgBox "Building Links: " &amp; vbCrLf &amp; strBuildingLinks, _ 
   vbInformation + vbOKOnly, "Building Links in Shared Workspace" 

```


## See also


#### Concepts


[SharedWorkspaceLink Object](sharedworkspacelink-object-office.md)
#### Other resources


[SharedWorkspaceLink Object Members](sharedworkspacelink-members-office.md)

