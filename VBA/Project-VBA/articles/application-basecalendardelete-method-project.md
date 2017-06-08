---
title: Application.BaseCalendarDelete Method (Project)
keywords: vbapj.chm619
f1_keywords:
- vbapj.chm619
ms.prod: project-server
api_name:
- Project.Application.BaseCalendarDelete
ms.assetid: f9583bd7-6ddb-7115-b7ca-c0e4e8b033e1
ms.date: 06/08/2017
---


# Application.BaseCalendarDelete Method (Project)

Deletes a base calendar.


## Syntax

 _expression_. **BaseCalendarDelete**( ** _Name_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|**String**. The name of the base calendar to delete.|

### Return Value

 **Boolean**


## Example

The following example deletes the base calendar entered by the user.


```vb
Sub DeleteCalendar() 
 
 Dim CalendarName As String 
 
 CalendarName = InputBox$("Enter name of base calendar to delete:") 
 BaseCalendarDelete Name:=CalendarName 
 
End Sub
```


