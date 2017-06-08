---
title: Project.ProjectGuideContent Property (Project)
keywords: vbapj.chm131091
f1_keywords:
- vbapj.chm131091
ms.prod: project-server
api_name:
- Project.Project.ProjectGuideContent
ms.assetid: 26ae4b6c-2541-b175-62d8-a4d4c015c6f5
ms.date: 06/08/2017
---


# Project.ProjectGuideContent Property (Project)

Gets or sets the name of the XML schema being used by the Project Guide. Read/write  **String**.


## Syntax

 _expression_. **ProjectGuideContent**

 _expression_ A variable that represents a **Project** object.


## Remarks


 **Note**  The Project Guide is deprecated in Project. Instead of the Project Guide, we recommend that you create task pane apps.

However, you can still use custom Project Guides and get the default Project Guide files from the Project SDK download. The Project Guide files are modified for access in a flat folder structure and to remove the  `gbui://` protocol ( **gbui** is the goal-based user interface protocol in Office Project 2007 and previous versions). All Project Guide settings must be made programmatically.

The default value of the  **ProjectGuideFunctionalLayoutPage** property is `gbui://mainpage.htm`, which does not work because Project does not implement the  `gbui://` protocol. The Project Programmability blog ( `http://blogs.msdn.com/project_programmability/`) includes articles that show how to use the Project Guide in a VBA macro and in an add-in that is developed with Visual C# in Microsoft Office development tools in Visual Studio 2010.


## Example

The following code sample changes the default content for the Project Guide to the XML file specified by the user. An input box prompts the user for the path and file name for custom Project Guide content.


 **Note**  Before running this macro, change  _path_ to an example path you would like to use, and change _filename_ to the name of an example file, such as custom.xml.


```vb
Sub UseCustomProjectGuide() 
   If Projects.Count = 0 Then 
      MsgBox "You must have at least one active project open." 
      Exit Sub 
   End If 
 
   Dim ProjectGuideURL As String 
   ProjectGuideURL = InputBox$(Prompt:="Enter the path and " _ 
      &; "file name of the XML file for custom Project " _ 
      &; "Guide content." &; Chr(13) _ 
      &; "For example, path \filename ") 
   If ProjectGuideURL = Empty Then 
      Exit Sub 
   Else 
      ActiveProject.ProjectGuideUseDefaultContent = False 
      ActiveProject.ProjectGuideContent = ProjectGuideURL 
      MsgBox Prompt:="The custom Project Guide content " _ 
         &; "defined in " &; ProjectGuideURL &; " is " _ 
         &; "now in use for the current project." 
   End If 
End Sub
```


