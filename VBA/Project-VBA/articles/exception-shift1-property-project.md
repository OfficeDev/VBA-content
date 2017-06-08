---
title: Exception.Shift1 Property (Project)
ms.prod: project-server
api_name:
- Project.Exception.Shift1
ms.assetid: 8b587014-c830-d346-4ba3-5add50f8e548
ms.date: 06/08/2017
---


# Exception.Shift1 Property (Project)

Gets a  **[Shift](shift-object-project.md)** object representing the first work shift in a calendar exception for a day, month, period, weekday, or throughout a year. Read-only **Shift**.


## Syntax

 _expression_. **Shift1**

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
