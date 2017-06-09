---
title: Task.ConstraintDate Property (Project)
keywords: vbapj.chm131666
f1_keywords:
- vbapj.chm131666
ms.prod: project-server
api_name:
- Project.Task.ConstraintDate
ms.assetid: 6985581b-82a1-6ab2-02ce-94d33e6d0336
ms.date: 06/08/2017
---


# Task.ConstraintDate Property (Project)

Gets or sets a constraint date for a task. Read/write  **Variant**.


## Syntax

 _expression_. **ConstraintDate**

 _expression_ A variable that represents a **Task** object.


## Example

The following example sets the constraint type to SNET and the constraint date to the current date for tasks in the active project with the default constraint of ASAP.


```vb
Sub SetConstraintDate() 
    Dim T As Task ' Task object used in For Each loop 
 
    For Each T In ActiveProject.Tasks 
        If T.ConstraintType = pjASAP Then 
            T.ConstraintType = pjSNET 
            T.ConstraintDate = ActiveProject.CurrentDate 
        End If 
    Next T 
End Sub
```


