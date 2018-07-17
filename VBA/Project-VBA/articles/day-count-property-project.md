---
title: Day.Count Property (Project)
ms.prod: project-server
api_name:
- Project.Day.Count
ms.assetid: 2f5c33fb-b744-6c50-5337-da693d93f28b
ms.date: 06/08/2017
---


# Day.Count Property (Project)

Gets the number of days in the  **Day** object, which is the value 1. Read-only **Integer**.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **Day** object.


## Example

The  **Count** property for the **Day** object is the value 1, as in the following example.


```vb
Debug.Print ActiveProject.Resources(1).Calendar.WeekDays(3).Count
```


