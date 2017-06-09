---
title: Application.FileSaveWorkspace Method (Project)
keywords: vbapj.chm108
f1_keywords:
- vbapj.chm108
ms.prod: project-server
api_name:
- Project.Application.FileSaveWorkspace
ms.assetid: f7c524e5-aa9e-e1a2-6f32-defb7cc23f04
ms.date: 06/08/2017
---


# Application.FileSaveWorkspace Method (Project)

Saves a list of open files and the current settings in the  **Options** dialog box.


## Syntax

 _expression_. **FileSaveWorkspace**( ** _Name_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the file to create. If  **Name** is omitted, Project prompts for the file name.|

### Return Value

 **Boolean**


## Example

The following example saves the workspace based upon the name of the first project file.


```vb
Sub SaveWorkspaceByProjectName() 
 
    Dim WSName As String 
 
    If InStr(Projects(1).Name, ".") Then 
        WSName = Left$(Projects(1).Name, Len(Projects(1).Name) - 1) &; "W" 
    Else 
        WSName = Projects(1).Name &; ".MPW" 
    End If 
 
    FileSaveWorkspace WSName 
End Sub
```


