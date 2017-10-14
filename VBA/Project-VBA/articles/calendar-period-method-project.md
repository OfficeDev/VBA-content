---
title: Calendar.Period Method (Project)
ms.prod: project-server
api_name:
- Project.Calendar.Period
ms.assetid: b717bcbe-654b-5791-2002-d65e2a96617f
ms.date: 06/08/2017
---


# Calendar.Period Method (Project)

Gets a  **[Period](period-object-project.md)** object representing a period of time in a calendar. Read-only **Period**.


## Syntax

 _expression_. **Period**( ** _Start_**, ** _Finish_** )

 _expression_ A variable that represents a **Calendar** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Required|**Variant**|The start date of the desired period.|
| _Finish_|Optional|**Variant**| The finish date of the desired period. The default value is the same date as Start.|

### Return Value

 **Period**


## Example

The following example sets a winter holiday for the active project.


```vb
Sub SetWinterHoliday() 
    ActiveProject.Calendar.Period("12/20/02", "12/31/02").Working = False 
 End Sub
```


