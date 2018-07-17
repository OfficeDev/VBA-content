---
title: Project.Index Property (Project)
ms.prod: project-server
api_name:
- Project.Project.Index
ms.assetid: 1213f55b-aca0-76ee-2e8a-2442a2c576e1
ms.date: 06/08/2017
---


# Project.Index Property (Project)

Gets the index of a  **Project** object in the containing **Projects** collection. Read-only **Variant**.


## Syntax

 _expression_. **Index**

 _expression_ A variable that represents a **Project** object.


## Example

If you put a Gantt chart in the same index of the  **Windows** collection for every open project, you can display a Gantt chart in one project and then use the **ActivateSameWindowInNextProject** macro to easily switch to the Gantt charts of the other open projects.


```vb
Sub ActivateSameWindowInNextProject() 
 
 ' Check for a next project. 
 If ActiveProject.Index = Application.Projects.Count Then 
 MsgBox("No more open projects") 
 ' Check for an equivalent window in the next project. 
 ElseIf ActiveProject.Windows.ActiveWindow.Index > Projects(ActiveProject.Index + 1).Windows.Count Then 
 MsgBox("No equivalent window in the next project") 
 ' If everything's okay, switch to the window in the next project. 
 Else 
 Projects(ActiveProject.Index + 1).Windows(ActiveWindow.Index).Activate 
 End If 
 
End Sub
```


