---
title: Month.Shift2 Property (Project)
ms.prod: project-server
api_name:
- Project.Month.Shift2
ms.assetid: 1499be01-4942-04b2-ff37-bbc0d49f9f68
ms.date: 06/08/2017
---


# Month.Shift2 Property (Project)

Gets a  **[Shift](shift-object-project.md)** object representing the second work shift in a month. Read-only **Shift**.


## Syntax

 _expression_. **Shift2**

 _expression_ A variable that represents a **Month** object.


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


