---
title: Task.ConstraintType Property (Project)
keywords: vbapj.chm131667
f1_keywords:
- vbapj.chm131667
ms.prod: project-server
api_name:
- Project.Task.ConstraintType
ms.assetid: cdcd6a0d-a996-646d-130e-1a5ed2c93705
ms.date: 06/08/2017
---


# Task.ConstraintType Property (Project)

Gets or sets a constraint type for a task. Read/write  **Variant**.


## Syntax

 _expression_. **ConstraintType**

 _expression_ A variable that represents a **Task** object.


## Remarks

The  **ConstraintType** property can be one of the **[PjConstraint](pjconstraint-enumeration-project.md)** constants.

If you set the  **ConstraintType** property to **pjFNET**, **pjFNLT**, **pjMFO**, **pjMSO**, **pjSNET**, or **pjSNLT**, Project uses the constraint date for the task. To set the constraint date, use the **[ConstraintDate](task-constraintdate-property-project.md)** property.


## Example

The following example changes the constraint type of tasks from MSO and MFO to SNET and FNLT.


```vb
Sub ChangeConstraintTypes() 
    Dim T As Task ' Task object used in For Each loop 
 
    For Each T In ActiveProject.Tasks 
        If T.ConstraintType = pjMSO Then 
            T.ConstraintType = pjSNET 
        ElseIf T.ConstraintType = pjMFO Then 
            T.ConstraintType = pjFNLT 
        End If 
    Next T 
End Sub
```


