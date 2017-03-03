---
title: SharedWorkspace.Name Property (Office)
keywords: vbaof11.chm276001
f1_keywords:
- vbaof11.chm276001
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SharedWorkspace.Name
ms.assetid: 2fec36b5-7455-6a0d-e381-fb21b0361d1e
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


```vb
Dim swsWorkspace As Office.SharedWorkspace 
    Dim strSWSInfo As String 
    Set swsWorkspace = ActiveWorkbook.SharedWorkspace 
    strSWSInfo = swsWorkspace.Name &; vbCrLf &; _ 
        " - URL: " &; swsWorkspace.URL &; vbCrLf &; _ 
        "The shared workspace contains " &; vbCrLf &; _ 
        " - Files: " &; swsWorkspace.Files.Count &; vbCrLf &; _ 
        " - Folders: " &; swsWorkspace.Folders.Count &; vbCrLf &; _ 
        " - Links: " &; swsWorkspace.Links.Count &; vbCrLf &; _ 
        " - Members: " &; swsWorkspace.Members.Count &; vbCrLf &; _ 
        " - Tasks: " &; swsWorkspace.Tasks.Count &; vbCrLf 
    MsgBox strSWSInfo, vbInformation + vbOKOnly, _ 
        "Shared Workspace Information" 
    Set swsWorkspace = Nothing
```


## See also


#### Concepts


[SharedWorkspace Object](sharedworkspace-object-office.md)

