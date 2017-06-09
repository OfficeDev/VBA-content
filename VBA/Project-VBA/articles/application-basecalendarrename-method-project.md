---
title: Application.BaseCalendarRename Method (Project)
keywords: vbapj.chm624
f1_keywords:
- vbapj.chm624
ms.prod: project-server
api_name:
- Project.Application.BaseCalendarRename
ms.assetid: e895c89f-1a29-0982-a88b-5af662215573
ms.date: 06/08/2017
---


# Application.BaseCalendarRename Method (Project)

Renames a base calendar.


## Syntax

 _expression_. **BaseCalendarRename**( ** _FromName_**, ** _ToName_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FromName_|Required|**String**|**String**. The name of the base calendar to rename.|
| _ToName_|Required|**String**|**String**. The new name of the base calendar.|

### Return Value

 **Boolean**


## Example

The following example changes the name of the base calendar from Night Shift to Third Shift.


```vb
Sub RenameNightShift() 
 BaseCalendarRename FromName:="Night Shift", ToName:="Third Shift" 
End Sub
```


