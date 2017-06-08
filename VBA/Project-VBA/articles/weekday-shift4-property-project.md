---
title: WeekDay.Shift4 Property (Project)
ms.prod: project-server
api_name:
- Project.WeekDay.Shift4
ms.assetid: 60b7fc25-b96d-6d85-672d-a5b6b6def322
ms.date: 06/08/2017
---


# WeekDay.Shift4 Property (Project)

Gets a  **[Shift](shift-object-project.md)** object representing the fourth work shift in a weekday. Read-only **Shift**.


## Syntax

 _expression_. **Shift4**

 _expression_ A variable that represents a **WeekDay** object.


## Example

The following example schedules a half-day of work on Fridays by creating an 8 A.M. to noon shift.


```vb
Sub HalfDayFridays() 
 
 With ActiveProject.Calendar.WeekDays(pjFriday) 
 .Shift1.Start = #8:00:00 AM# 
 .Shift1.Finish = #12:00:00 PM# 
 .Shift2.Clear 
 .Shift3.Clear 
 End With 
 
End Sub
```


