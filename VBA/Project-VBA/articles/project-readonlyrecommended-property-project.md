---
title: Project.ReadOnlyRecommended Property (Project)
ms.prod: project-server
api_name:
- Project.Project.ReadOnlyRecommended
ms.assetid: f35003bc-97fb-3acd-f629-7bb8addc5261
ms.date: 06/08/2017
---


# Project.ReadOnlyRecommended Property (Project)

 **True** if the project should be opened with read-only access. Read-only **Boolean**.


## Syntax

 _expression_. **ReadOnlyRecommended**

 _expression_ A variable that represents a **Project** object.


## Remarks

To change the value of the  **ReadOnlyRecommended** property, use the **[FileSaveAs](application-filesaveas-method-project.md)** method with the ReadOnly argument set to **True**.


## Example

The following example displays the recommended access type for the active project.


```vb
Sub DisplayAccessType() 
    If ActiveProject.ReadOnlyRecommended Then 
        MsgBox "Read-only access is recommended for this project." 
    ElseIf ActiveProject.ReadOnly Then 
        MsgBox "This project may only be opened read-only." 
    Else 
        MsgBox "Read/write access is allowed for this project." 
    End If 
End Sub
```


