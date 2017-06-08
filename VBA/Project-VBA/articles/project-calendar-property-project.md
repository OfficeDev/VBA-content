---
title: Project.Calendar Property (Project)
ms.prod: project-server
api_name:
- Project.Project.Calendar
ms.assetid: 0496a31e-7469-57e0-7675-ac9c6677f992
ms.date: 06/08/2017
---


# Project.Calendar Property (Project)

Gets a  **[Calendar](calendar-object-project.md)** object representing a calendar for the project. Read-only **Calendar**.


## Syntax

 _expression_. **Calendar**

 _expression_ A variable that represents a **Project** object.


## Example

The following example resets the calendar for the active project.


```vb
Sub ResetActiveProjectCalendar() 
 ActiveProject.Calendar.Reset 
End Sub
```


