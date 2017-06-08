---
title: WeekDays.Count Property (Project)
ms.prod: project-server
api_name:
- Project.WeekDays.Count
ms.assetid: 6343346c-dbfc-b36b-eaf4-ddcc2e6f745d
ms.date: 06/08/2017
---


# WeekDays.Count Property (Project)

Gets the number of items in the  **WeekDays** collection. Read-only **Integer**.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **WeekDays** object.


## Example

The following example shows there are seven days in the week for the calendar of the specified resource.


```vb
Debug.Print ActiveProject.Resources(1).Calendar.WorkWeeks(1).WeekDays.Count
```


## See also


#### Concepts


[WeekDays Collection Object](weekdays-object-project.md)
