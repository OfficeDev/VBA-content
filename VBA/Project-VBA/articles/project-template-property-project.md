---
title: Project.Template Property (Project)
keywords: vbapj.chm132720
f1_keywords:
- vbapj.chm132720
ms.prod: project-server
api_name:
- Project.Project.Template
ms.assetid: 8f73cf7a-e900-2951-6491-edc0ef78c0f5
ms.date: 06/08/2017
---


# Project.Template Property (Project)

Gets the name of the template associated with a project. Read-only  **String**.


## Syntax

 _expression_. **Template**

 _expression_ A variable that represents a **Project** object.


## Remarks

If the project was not created from a template, the  **Template** property returns an empty string ("").


## Example

The following example creates a new project based on the template of the active project, if the active project was previously created from a Project template file (.mpt).


```vb
Sub CreateNewProject() 
    FileOpen ActiveProject.Template &; ".mpt" 
End Sub
```


