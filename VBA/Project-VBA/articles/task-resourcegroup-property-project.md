---
title: Task.ResourceGroup Property (Project)
ms.prod: project-server
api_name:
- Project.Task.ResourceGroup
ms.assetid: 3ff88223-3b9c-cf5a-559c-7e41d7ed2e33
ms.date: 06/08/2017
---


# Task.ResourceGroup Property (Project)

Gets the names of groups associated with the resources assigned to a task, separated by the list separator. Read-only  **String**.


## Syntax

 _expression_. **ResourceGroup**

 _expression_ A variable that represents a **Task** object.


## Remarks

For example, if Bob's group is "Writers" and Greg's group is "Editors", and Greg and Bob are assigned to the same task, then the  **ResourceGroup** property for that task returns "Writers,Editors". This example assumes that the list separator character is the comma (,). The list separator character can be set with the **ListSeparator** property.


