---
title: Exception.Shift5 Property (Project)
ms.prod: project-server
api_name:
- Project.Exception.Shift5
ms.assetid: 1275285a-3471-08bd-12b6-d37e60e4d9be
ms.date: 06/08/2017
---


# Exception.Shift5 Property (Project)

Gets a  **[Shift](shift-object-project.md)** object representing the fifth work shift in a calendar exception for a day, month, period, weekday, or throughout a year. Read-only **Shift**.


## Syntax

 _expression_. **Shift5**

 _expression_ A variable that represents an **Exception** object.


## Example

The following example schedules a half-day of work on Fridays by creating a shift from 8 A.M. to noon.


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


## See also


#### Concepts


[Exception Object](exception-object-project.md)
