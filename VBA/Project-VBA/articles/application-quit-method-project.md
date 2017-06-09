---
title: Application.Quit Method (Project)
keywords: vbapj.chm131252
f1_keywords:
- vbapj.chm131252
ms.prod: project-server
api_name:
- Project.Application.Quit
ms.assetid: 0aaba635-6d6a-c4a3-fab3-03451659021b
ms.date: 06/08/2017
---


# Application.Quit Method (Project)

Exits Microsoft Project.


## Syntax

 _expression_. **Quit**( ** _SaveChanges_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SaveChanges_|Optional|**Long**|Specifies whether Project saves changes before quitting. Can be one of the following  **[PjSaveType](pjsavetype-enumeration-project.md)** constants: **pjDoNotSave**, **pjSave**, or **pjPromptSave**. The default is **pjPromptSave** for new project files and projects that have changed since the last save.|

## Example

The following example saves all open projects and then exits Project.


```vb
Sub SaveChangesAndQuit() 
 Quit SaveChanges:=pjSave 
End Sub
```


