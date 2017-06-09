---
title: Project.ProjectGuideSaveBuffer Property (Project)
keywords: vbapj.chm131087
f1_keywords:
- vbapj.chm131087
ms.prod: project-server
api_name:
- Project.Project.ProjectGuideSaveBuffer
ms.assetid: 0e2ef078-6667-fd17-f240-4c847c88dbeb
ms.date: 06/08/2017
---


# Project.ProjectGuideSaveBuffer Property (Project)

Gets or sets an XML string representing the save buffer of the Project Guide. Read/write  **String**.


## Syntax

 _expression_. **ProjectGuideSaveBuffer**

 _expression_ A variable that represents a **Project** object.


## Remarks


 **Note**  The Project Guide is deprecated in Project. Instead of the Project Guide, we recommend that you create task pane apps.

However, you can still use custom Project Guides and get the default Project Guide files from the Project SDK download. The Project Guide files are modified for access in a flat folder structure and to remove the  `gbui://` protocol ( **gbui** is the goal-based user interface protocol in Office Project 2007 and previous versions). All Project Guide settings must be made programmatically.

The default value of the  **ProjectGuideFunctionalLayoutPage** property is `gbui://mainpage.htm`, which does not work because Project does not implement the  `gbui://` protocol. The Project Programmability blog ( `http://blogs.msdn.com/project_programmability/`) includes articles that show how to use the Project Guide in a VBA macro and in an add-in that is developed with Visual C# in Microsoft Office development tools in Visual Studio 2010.


