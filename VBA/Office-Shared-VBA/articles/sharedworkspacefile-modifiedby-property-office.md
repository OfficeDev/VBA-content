---
title: SharedWorkspaceFile.ModifiedBy Property (Office)
keywords: vbaof11.chm266004
f1_keywords:
- vbaof11.chm266004
ms.prod: office
api_name:
- Office.SharedWorkspaceFile.ModifiedBy
ms.assetid: d6533854-ddd9-3a41-b74b-94f282779236
ms.date: 06/08/2017
---


# SharedWorkspaceFile.ModifiedBy Property (Office)

Gets the name of the user who last modified the object. Read-only.


## Syntax

 _expression_. **ModifiedBy**

 _expression_ A variable that represents a **SharedWorkspaceFile** object.


### Return Value

String


## Remarks

For shared workspace objects, the  **ModifiedBy** property returns the display name stored in the **Name** property of the **SharedWorkspaceMember** object.


## Example

The following example lists the files in a shared workspace site that were last modified by users other than the creator of the workspace site.


```
 Dim swsFile As Office.SharedWorkspaceFile 
 Dim swsOwner As Office.SharedWorkspaceMember 
 Dim strMemberFiles As String 
 Set swsOwner = ActiveWorkbook.SharedWorkspace.Members(1) 
 For Each swsFile In ActiveWorkbook.SharedWorkspace.Files 
 If swsFile.ModifiedBy <> swsOwner.Name Then 
 strMemberFiles = strMemberFiles &amp; swsFile.URL &amp; vbCrLf 
 End If 
 Next 
 MsgBox "These files were last modified by other users:" &amp; _ 
 vbCrLf &amp; strMemberFiles, _ 
 vbInformation + vbOKOnly, "Files Modified by Other Users" 
 Set swsOwner = Nothing 
 Set swsFile = Nothing 

```


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## See also


#### Concepts


[SharedWorkspaceFile Object](sharedworkspacefile-object-office.md)
#### Other resources


[SharedWorkspaceFile Object Members](sharedworkspacefile-members-office.md)

