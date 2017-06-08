---
title: Application.FileExit Method (Project)
keywords: vbapj.chm114
f1_keywords:
- vbapj.chm114
ms.prod: project-server
api_name:
- Project.Application.FileExit
ms.assetid: a69bc574-dcc3-3710-c705-0566fcf10235
ms.date: 06/08/2017
---


# Application.FileExit Method (Project)

Quits Project.


## Syntax

 _expression_. **FileExit**( ** _Save_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Save_|Optional|**Long**|Can be one of the [PjSaveType](pjsavetype-enumeration-project.md) constants. The default value is **pjPromptSave** for new project files and projects that have changed since the last save.|

### Return Value

 **Boolean**


## Example

The following example saves and closes the active project, and then exits the Project application.


```vb
Sub SaveAndCloseActiveProject() 
    FileExit pjSave 
End Sub
```


