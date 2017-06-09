---
title: Application.GlobalBaseCalendars Property (Project)
keywords: vbapj.chm132295
f1_keywords:
- vbapj.chm132295
ms.prod: project-server
api_name:
- Project.Application.GlobalBaseCalendars
ms.assetid: 98a498f9-e040-9b00-e84a-806a8a17a181
ms.date: 06/08/2017
---


# Application.GlobalBaseCalendars Property (Project)

Gets or sets a  **[Calendars](calendar-object-project.md)** collection representing the base calendars of the Global.mpt file. Read/write **Calendars**.


## Syntax

 _expression_. **GlobalBaseCalendars**

 _expression_ A variable that represents an **Application** object.


## Remarks

 To add a calendar to the enterprise global template, first create a local calendar, and then add the local calendar to the enterprise global template with the **MakeLocalCalendarEnterprise** method.

To enable creating local base calendars in an enterprise project, check  **Allow projects to use local base calendars** on the Additional Server Settings page in Project Web App.


## Example

The following example creates a local base calendar and then imports the calendar to the enterprise global template.


 **Note**  The  **GlobalBaseCalendars** property is the collection of calendars in the local Global.mpt file, not in the enterprise global template.


```vb
Sub CreateEGlobalCalendar() 
    Dim globalCalendar As Calendar 
 
    BaseCalendarCreate Name:="NewBaseCalendar" 
    MakeLocalCalendarEnterprise OldName:="NewBaseCalendar", NewName:="NewBaseCalendar" 
 
    Debug.Print "Number of calendars in Global.mpt: " &; GlobalBaseCalendars.Count 
 
    For Each globalCalendar In GlobalBaseCalendars 
        Debug.Print globalCalendar.Name 
    Next globalCalendar 
End Sub
```


