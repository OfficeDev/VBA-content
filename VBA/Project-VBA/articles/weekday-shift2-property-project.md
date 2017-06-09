---
title: WeekDay.Shift2 Property (Project)
ms.prod: project-server
api_name:
- Project.WeekDay.Shift2
ms.assetid: cc686169-6b7d-099a-6602-9381418fa0de
ms.date: 06/08/2017
---


# WeekDay.Shift2 Property (Project)

Gets a  **[Shift](shift-object-project.md)** object representing the second work shift in a weekday. Read-only **Shift**.


## Syntax

 _expression_. **Shift2**

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


