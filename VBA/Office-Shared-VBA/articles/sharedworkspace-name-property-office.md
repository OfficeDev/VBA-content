---
title: SharedWorkspace.Name Property (Office)
keywords: vbaof11.chm276001
f1_keywords:
- vbaof11.chm276001
ms.prod: office
api_name:
- Office.SharedWorkspace.Name
ms.assetid: 2fec36b5-7455-6a0d-e381-fb21b0361d1e
ms.date: 06/08/2017
---


# SharedWorkspace.Name Property (Office)

Gets or sets the display name of the shared workspace site. Read/write.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents a **SharedWorkspace** object.


### Return Value

String


## Example

The following example displays the properties of the shared workspace to which the active document is connected.


```
Dim swsWorkspace As Office.SharedWorkspace 
    Dim strSWSInfo As String 
    Set swsWorkspace = ActiveWorkbook.SharedWorkspace 
    strSWSInfo = swsWorkspace.Name &amp; vbCrLf &amp; _ 
        " - URL: " &amp; swsWorkspace.URL &amp; vbCrLf &amp; _ 
        "The shared workspace contains " &amp; vbCrLf &amp; _ 
        " - Files: " &amp; swsWorkspace.Files.Count &amp; vbCrLf &amp; _ 
        " - Folders: " &amp; swsWorkspace.Folders.Count &amp; vbCrLf &amp; _ 
        " - Links: " &amp; swsWorkspace.Links.Count &amp; vbCrLf &amp; _ 
        " - Members: " &amp; swsWorkspace.Members.Count &amp; vbCrLf &amp; _ 
        " - Tasks: " &amp; swsWorkspace.Tasks.Count &amp; vbCrLf 
    MsgBox strSWSInfo, vbInformation + vbOKOnly, _ 
        "Shared Workspace Information" 
    Set swsWorkspace = Nothing
```


## See also


#### Concepts


[SharedWorkspace Object](sharedworkspace-object-office.md)
#### Other resources


[SharedWorkspace Object Members](sharedworkspace-members-office.md)

