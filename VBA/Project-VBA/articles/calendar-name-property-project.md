---
title: Calendar.Name Property (Project)
ms.prod: PROJECTSERVER
api_name:
- Project.Calendar.Name
ms.assetid: e437e29c-ed61-c83a-53b7-8a0d1cb7cb4e
---


# Calendar.Name Property (Project)

Gets the name of a  **Calendar** object. Read-only **String**.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents a **Calendar** object.


## Remarks

For a code example that uses the  **Task** object, see **[Name](task-name-property-project.md)**.


## Example

 **Name** is the default property for the **Calendar** object. The following example prints the name of the default calendar for the active project.


```vb
Debug.Print ActiveProject.Calendar
```


