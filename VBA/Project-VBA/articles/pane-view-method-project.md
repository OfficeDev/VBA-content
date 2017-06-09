---
title: Pane.View Method (Project)
ms.prod: project-server
api_name:
- Project.Pane.View
ms.assetid: a29aa7d4-e712-bbf4-96dd-e0fdeab70ba2
ms.date: 06/08/2017
---


# Pane.View Method (Project)

Returns the active  **View** object.


## Syntax

 _expression_. **View**

 _expression_ A variable that represents a **Pane** object.


### Return Value

 **View**


## Example

The following statement prints the name of the view in the  **Immediate** window in the VBE. For example, if the Team Planner view is active, the statement prints "Team Plannner".


```vb
Debug.Print ActiveWindow.ActivePane.View.Name
```


