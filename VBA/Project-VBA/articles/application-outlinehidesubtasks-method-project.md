---
title: Application.OutlineHideSubTasks Method (Project)
keywords: vbapj.chm2020
f1_keywords:
- vbapj.chm2020
ms.prod: project-server
api_name:
- Project.Application.OutlineHideSubTasks
ms.assetid: 79e79b71-aa4d-eb17-7f27-96d4dd382547
ms.date: 06/08/2017
---


# Application.OutlineHideSubTasks Method (Project)

Hides the subtasks of the selected task or tasks.


## Syntax

 _expression_. **OutlineHideSubTasks**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **Boolean**


## Example

The following example collapses the entire outline of the first task.


```vb
Sub OutlineHideAllSubtasks() 
 ActiveProject.Tasks(1).OutlineHideSubtasks 
End Sub
```


