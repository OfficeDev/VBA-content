---
title: Application.ActiveProject Property (Project)
keywords: vbapj.chm131377
f1_keywords:
- vbapj.chm131377
ms.prod: project-server
api_name:
- Project.Application.ActiveProject
ms.assetid: 07844166-ca9b-15eb-a5e2-6f00a7c0a030
ms.date: 06/08/2017
---


# Application.ActiveProject Property (Project)

Gets a  **[Project](project-object-project.md)** object that represents the active project. Read-only **Project**.


## Syntax

 _expression_. **ActiveProject**

 _expression_ A variable that represents an **Application** object.


## Example

The following example adds the date and time to the  **Comments** field in the project **Properties** dialog box and then saves the project.


```vb
Sub SaveAndNoteTime() 
 ActiveProject.ProjectNotes = ActiveProject.ProjectNotes &; vbCrLf _ 
 &; "This project was last saved on " &; Date$ &; " at " &; Time$ &; "." 
 FileSave 
End Sub
```


