---
title: Application.Caption Property (Project)
ms.prod: project-server
api_name:
- Project.Application.Caption
ms.assetid: e43c55ea-d239-a6e5-42ce-35da5b47aa01
ms.date: 06/08/2017
---


# Application.Caption Property (Project)

Gets or sets the text in the title bar of the main window. Read/write  **String**.


## Syntax

 _expression_. **Caption**

 _expression_ A variable that represents an **Application** object.


## Remarks

When the active window is maximized, the title bar displays the caption for both the main and active windows, separating the captions with a hyphen. For example, if the caption for the main window is "Microsoft Project" and the caption for the active window is "Project1.mpp", the title bar displays "Project1.mpp - Microsoft Project" when the active window is maximized.

If you set the  **Caption** property to **Empty**, the title bar displays the default caption. The default caption for the main window is "Microsoft Project".



In a project with one window, the default caption for the window is the file name of the project. In a project that has multiple windows open, the default caption for each window is  _name_: _n,_ where _name_ is the file name of the project and _n_ is a unique number for the window. For example, if the second window of the project "Project1" is active, the default title bar displays "Project1.mpp:2 - Microsoft Project"


## Example

The following example prompts the user to change the caption for the active window.


```vb
Sub ChangeWindowCaption() 
 
 Dim Entry As String ' Caption entered by user 
 
 ' Prompt user for a new caption. 
 Entry = InputBox$("Enter a new caption for the active window (enter 'reset' to set the caption to its default).") 
 
 ' If user chooses the Cancel button, exit Sub procedure. 
 If Entry = Empty Then Exit Sub 
 
 ' Set or reset the caption. 
 If Entry = "reset" Then 
 ActiveWindow.Caption = Empty 
 Else 
 ActiveWindow.Caption = Entry 
 End If 
 
End Sub
```


