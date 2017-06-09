---
title: Application.BaseCalendarCreate Method (Project)
keywords: vbapj.chm618
f1_keywords:
- vbapj.chm618
ms.prod: project-server
api_name:
- Project.Application.BaseCalendarCreate
ms.assetid: c9c92dff-255a-041b-c18d-49d6d75884e3
ms.date: 06/08/2017
---


# Application.BaseCalendarCreate Method (Project)

Creates a base calendar.


## Syntax

 _expression_. **BaseCalendarCreate**( ** _Name_**, ** _FromName_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the base calendar to create.|
| _FromName_|Optional|**String**|The name of the base calendar to copy.|

### Return Value

 **Boolean**


## Remarks

To create a local calendar when Project Professional is logged on to Project Server, you must check  **Allow projects to use local base calendars** on the Additional Server Settings page in Project Web Access. Restart Project Professional after changing the setting in Project Web Access.


## Example

The following example creates a new base calendar called "Base Holiday Calendar."


```vb
Sub CreateHolidayCalendar() 
 BaseCalendarCreate Name:="Base Holiday Calendar" 
End Sub
```


