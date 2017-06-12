---
title: Years.Count Property (Project)
ms.prod: project-server
api_name:
- Project.Years.Count
ms.assetid: 6a65ff7b-55ca-31e0-0edd-c2f75cb9fc74
ms.date: 06/08/2017
---


# Years.Count Property (Project)

Gets the number of items in the  **Years** collection. Read-only **Integer**.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **Years** object.


## Remarks

The following statement prints 166 in the  **Immediate** pane of the VBE. The value is the number of years from 1984 to and including 2149.


```
Print ActiveProject.Calendar.Years.Count
```

Use of the  **Count** property in most collection objects is similar. For an example that uses the **Years** collection, see[Years Object](years-object-project.md).


## See also


#### Concepts


[Years Collection Object](years-object-project.md)
