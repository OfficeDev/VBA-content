---
title: Project.CodeName Property (Project)
ms.prod: project-server
api_name:
- Project.Project.CodeName
ms.assetid: 78c63fe2-30ad-bee7-1901-2fa0e293c7b4
ms.date: 06/08/2017
---


# Project.CodeName Property (Project)

Gets the code name for the project. Read-only  **String**.


## Syntax

 _expression_. **CodeName**

 _expression_ A variable that represents a **Project** object.


## Remarks

The code name is the name of the module that stores event macros (and other macros you may have defined) for a project. The default name for the module is "ThisProject"; you can view it in the  **Project** window in the Visual Basic Editor.

Changing the project name doesn't change the code name, and changing the code name (using the  **Properties** window in the Visual Basic Editor) doesn't change the project name.


