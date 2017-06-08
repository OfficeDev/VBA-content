---
title: Project.CommandBars Property (Project)
ms.prod: project-server
api_name:
- Project.Project.CommandBars
ms.assetid: 8b987a76-0aa4-537b-871b-ad36338b2b4e
ms.date: 06/08/2017
---


# Project.CommandBars Property (Project)

Gets a  **CommandBars** collection that represents all the command bars in the project. Read-only **CommandBars**.


## Syntax

 _expression_. **CommandBars**

 _expression_ A variable that represents a **Project** object.


## Remarks

For more information, see the  **CommandBars** object in the Office Developer Reference.


## Example

The following example lists all command bars in the project that are not currently visible.


```vb
Sub ListCommandBars() 
    Dim Bar As CommandBar 
     
    For Each Bar In ActiveProject.CommandBars 
        If Not Bar.Visible Then Debug.Print Bar.Name 
    Next 
End Sub
```


