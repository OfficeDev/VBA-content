---
title: Project.ID Property (Project)
ms.prod: project-server
api_name:
- Project.Project.ID
ms.assetid: d21541b3-d6ff-546e-8207-48b8cd180d2c
ms.date: 06/08/2017
---


# Project.ID Property (Project)

Gets the identification number of a project. Read-only  **Long**.


## Syntax

 _expression_. **ID**

 _expression_ A variable that represents a **Project** object.


## Remarks

The  **ID** property of a project can change when the project is closed and then opened again after other projects are opened. Use the **UniqueID** property if you want a constant reference to a project.


