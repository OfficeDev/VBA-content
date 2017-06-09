---
title: Calendar.WeekDays Property (Project)
keywords: vbapj.chm132819
f1_keywords:
- vbapj.chm132819
ms.prod: project-server
api_name:
- Project.Calendar.WeekDays
ms.assetid: 4495a739-156b-8cda-d3d0-acbc56b767ff
ms.date: 06/08/2017
---


# Calendar.WeekDays Property (Project)

Gets a  **[Weekdays](weekday-object-project.md)** collection representing the weekdays in the calendar. Read-only **Weekdays**.


## Syntax

 _expression_. **WeekDays**

 _expression_ A variable that represents a **Calendar** object.


## Example

The following example makes Friday a nonworking day in the calendar for the active project.


```vb
Sub MakeFridaysNonworking() 
 ActiveProject.Calendar.Weekdays(pjFriday).Working = False 
End Sub
```


