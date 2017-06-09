---
title: Application.ChangeWorkingTimeEx Method (Project)
keywords: vbapj.chm625
f1_keywords:
- vbapj.chm625
ms.prod: project-server
api_name:
- Project.Application.ChangeWorkingTimeEx
ms.assetid: 4608fdab-0b39-9918-522a-71d502ba7e3a
ms.date: 06/08/2017
---


# Application.ChangeWorkingTimeEx Method (Project)

Displays the  **Change Working Time** dialog box, which prompts the user to change a calendar.


## Syntax

 _expression_. **ChangeWorkingTimeEx**( ** _CalendarName_**, ** _Locked_**, ** _SelectedDate_**, ** _ProjectName_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CalendarName_|Optional|**String**|The name of the calendar.|
| _Locked_|Optional|**Boolean**|**True** if Project disables the **New** and **Options** buttons in the **Change Working Time** dialog box. The default value is **False**.|
| _SelectedDate_|Optional|**Variant**||
| _ProjectName_|Optional|**Variant**|Name of the project to change. The default is the active project.|

### Return Value

 **Boolean**


## Remarks

The  **ChangeWorkingTime** method has the same effect as the **Change Working Time** command on the **Project** tab in the Project Ribbon.


