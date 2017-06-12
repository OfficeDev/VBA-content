---
title: Project.FullName Property (Project)
ms.prod: project-server
api_name:
- Project.Project.FullName
ms.assetid: ae8cea25-f365-d8ae-e119-929a61a9c110
ms.date: 06/08/2017
---


# Project.FullName Property (Project)

Gets the path and file name of a project. Read-only  **String**.


## Syntax

 _expression_. **FullName**

 _expression_ A variable that represents a **Project** object.


## Remarks

The  **FullName** property returns the project name (as seen in the title bar) for an unsaved project.


## Example

The following example prompts the user for the full name of a file and then closes the file, saving it if it has changed.


```vb
Sub CloseFile() 
 Dim P As Project ' Project object used in For Each loop 
 Dim FileName As String ' Full name of a file 
 
 ' Prompt user for the full name of a file. 
 FileName = InputBox$("Close which file? Include its path: ") 
 
 ' Search the open projects for the file. 
 For Each P In Application.Projects 
 
 ' If the file is found, close it. 
 If P.FullName = FileName Then 
 P.Activate 
 FileClose pjSave 
 Exit Sub 
 End If 
 Next P 
 
 ' Inform user if the file is not found. 
 MsgBox ("Could not find the file " &; FileName &; ".") 
 
End Sub
```


