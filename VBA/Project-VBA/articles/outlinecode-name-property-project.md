---
title: OutlineCode.Name Property (Project)
ms.prod: project-server
api_name:
- Project.OutlineCode.Name
ms.assetid: b4814e58-2efd-18aa-4018-eb883fc64afa
ms.date: 06/08/2017
---


# OutlineCode.Name Property (Project)

Gets the name of the  **OutlineCode** object. Read/write **String**.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents an **OutlineCode** object.


## Remarks

For a code example that uses the  **Task** object, see **[Name](task-name-property-project.md)**.


## Example

 **Name** is the default property for the **OutlineCode** object. If the first task outline code for the active project is defined, the following example prints the name of the outline code.


```vb
Debug.Print ActiveProject.OutlineCodes(1)
```


