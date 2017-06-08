---
title: Application.IsUndoingOrRedoing Method (Project)
ms.prod: project-server
api_name:
- Project.Application.IsUndoingOrRedoing
ms.assetid: e0e5ddc7-aa22-0d43-1de6-83a260d57608
ms.date: 06/08/2017
---


# Application.IsUndoingOrRedoing Method (Project)

Indicates whether Project is currently executing an undo or redo action.


## Syntax

 _expression_. **IsUndoingOrRedoing**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **Boolean**


## Remarks

 Use the **[Application.OnUndoOrRedo ](application-onundoorredo-event-project.md)** event to listen for specific undo or redo actions.


