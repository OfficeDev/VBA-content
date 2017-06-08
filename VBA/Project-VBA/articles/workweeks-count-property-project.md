---
title: WorkWeeks.Count Property (Project)
ms.prod: project-server
api_name:
- Project.WorkWeeks.Count
ms.assetid: d8360e75-7dbe-955b-dd95-20fb3bf465e3
ms.date: 06/08/2017
---


# WorkWeeks.Count Property (Project)

Gets the number of items in the  **WorkWeeks** collection. Read-only **Long**.


## Syntax

 _expression_. **Count**

 _expression_ An expression that returns a **WorkWeeks** object.


## Example

The following example shows the number of custom work weeks defined in the calendar for the first resource in the active project.


```vb
Debug.Print ActiveProject.Resources(1).Calendar.WorkWeeks.Count
```


## See also


#### Concepts


[WorkWeeks Collection Object](workweeks-object-project.md)
