---
title: Project.UniqueID Property (Project)
ms.prod: project-server
api_name:
- Project.Project.UniqueID
ms.assetid: b49c0065-4b74-4e8e-48fa-9cf80bfc6e34
ms.date: 06/08/2017
---


# Project.UniqueID Property (Project)

Gets the unique identification number of the project, which is actually the  **UniqueID** value of the project summary task. Read-only **Long**.


## Syntax

 _expression_. **UniqueID**

 _expression_ A variable that represents a **Project** object.


## Remarks

The  **UniqueID** value of a project is 0, which is the value for the project summary task. If there are multiple projects open, the **ID** property of each project represents the order in which the projects are opened (1, 2, 3, and so forth), but the **UniqueID** of each project is 0.

If a project contains a subproject, and the master project is the only one open, the  `Application.Projects.Count` statement returns the value 2. The `Application.Projects(2).ID` value is 2, but the `Application.Projects(2).UniqueID` value is still 0.


