---
title: Application.FileCloseEx Method (Project)
keywords: vbapj.chm103
f1_keywords:
- vbapj.chm103
ms.prod: project-server
api_name:
- Project.Application.FileCloseEx
ms.assetid: 56e6eec6-6031-312b-fba5-50db7b43f0b1
ms.date: 06/08/2017
---


# Application.FileCloseEx Method (Project)

Closes the active project.


## Syntax

 _expression_. **FileCloseEx**( ** _Save_**, ** _NoAuto_**, ** _CheckIn_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Save_|Optional|**Long**|Can be one of the following  **PjSave** constants: **pjDoNotSave**, **pjSave**, or **pjPromptSave**. The default value is **pjPromptSave** for new project files and projects that have changed since the last save.|
| _NoAuto_|Optional|**Boolean**|**True** if an **Auto_Close** macro is not run and the **Close** event is not raised. The default value is **False**.|
| _CheckIn_|Optional|**Variant**|**True** if file is checked in after closing. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

The CheckIn parameter can accept the value  **True**, **False**, 0, 1, "Yes", or "No".


## Example

The following example saves and closes the active project.


```vb
Sub SaveAndCloseActiveProject() 
 FileCloseEx pjSave 
End Sub
```


