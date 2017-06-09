---
title: Project.LastSavedBy Property (Project)
ms.prod: project-server
api_name:
- Project.Project.LastSavedBy
ms.assetid: bc0d7330-1d58-5cc4-998c-b070450a7832
ms.date: 06/08/2017
---


# Project.LastSavedBy Property (Project)

Gets the name of the user who last saved a project. Read-only  **String**.


## Syntax

 _expression_. **LastSavedBy**

 _expression_ A variable that represents a **Project** object.


## Example

The following example adds the date the active project was last saved and the name of the user who last saved it to the notes of the active project


```vb
Sub AddSaveInfoToNotes() 
 ActiveProject.ProjectNotes = ActiveProject.ProjectNotes &; vbCrLf &; "This project was last saved on " &; CStr(ActiveProject.LastSaveDate) &; " by " &; ActiveProject.LastSavedBy &; "." 
End Sub
```


