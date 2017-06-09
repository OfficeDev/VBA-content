---
title: WeekDay.Shift3 Property (Project)
ms.prod: project-server
api_name:
- Project.WeekDay.Shift3
ms.assetid: c09fde08-3f8d-71e8-5c5d-f0ebbb0069ce
ms.date: 06/08/2017
---


# WeekDay.Shift3 Property (Project)

Gets a  **[Shift](shift-object-project.md)** object representing the third work shift in a weekday. Read-only **Shift**.


## Syntax

 _expression_. **Shift3**

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


