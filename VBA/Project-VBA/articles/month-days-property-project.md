---
title: Month.Days Property (Project)
ms.prod: project-server
api_name:
- Project.Month.Days
ms.assetid: 86572272-1a5f-2c86-2111-e41f39f4c1e6
ms.date: 06/08/2017
---


# Month.Days Property (Project)

Gets a  **[Days](day-object-project.md)** collection representing the days in the month. Read-only **Days**.


## Syntax

 _expression_. **Days**

 _expression_ A variable that represents a **Month** object.


## Example

The following example makes January 1 of every year a nonworking day.


```vb
Sub NewYearsDayOff() 
 
 Dim Y As Year 
 
 For Each Y In ActiveProject.Calendar.Years 
 Y.Months(pjJanuary).Days(1).Working = False 
 Next Y 
 
End Sub
```


