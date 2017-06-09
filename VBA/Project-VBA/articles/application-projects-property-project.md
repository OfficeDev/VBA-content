---
title: Application.Projects Property (Project)
ms.prod: project-server
api_name:
- Project.Application.Projects
ms.assetid: 792b7334-a424-abe1-287e-285d3ab362c7
ms.date: 06/08/2017
---


# Application.Projects Property (Project)

Gets a  **[Projects](projects-object-project.md)** collection representing the open projects. Read-only **Projects**.


## Syntax

 _expression_. **Projects**

 _expression_ A variable that represents an **Application** object.


## Remarks

To see the  **Project Properties** dialog box, choose the **FILE** tab to show the **Backstage** view. On the **Info** tab, select the **Project Information** drop-down menu, and then choose **Advanced Properties**.


## Example

The following example adds the date and time to the  **Comments** field in the project **Properties** dialog box, and then saves the project.


```vb
Sub SaveAndNoteTime() 
    Projects(1).ProjectNotes = Projects(1).ProjectNotes &; vbCrLf _ 
        &; "This project was last saved on " _ 
        &; Date$ &; " at " &; Time$ &; "." 
    FileSave 
End Sub
```


