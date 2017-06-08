---
title: Day.Shift3 Property (Project)
ms.prod: project-server
api_name:
- Project.Day.Shift3
ms.assetid: c8a70ddf-ef14-3388-3ddb-9e0e35d8a665
ms.date: 06/08/2017
---


# Day.Shift3 Property (Project)

Gets a  **[Shift](shift-object-project.md)** object representing the third work shift in a day. Read-only **Shift**.


## Syntax

 _expression_. **Shift3**

 _expression_ A variable that represents a **Day** object.


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


