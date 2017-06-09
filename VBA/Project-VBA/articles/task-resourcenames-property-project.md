---
title: Task.ResourceNames Property (Project)
keywords: vbapj.chm132571
f1_keywords:
- vbapj.chm132571
ms.prod: project-server
api_name:
- Project.Task.ResourceNames
ms.assetid: 0c933d60-42bf-ece6-fa37-da5181a56944
ms.date: 06/08/2017
---


# Task.ResourceNames Property (Project)

Gets or sets the names of the resources assigned to a task. Read/write  **String**.


## Syntax

 _expression_. **ResourceNames**

 _expression_ A variable that represents a **Task** object.


## Remarks

For a task with more than one resource, the  **ResourceNames** property returns the names of the resources, separated by the list separator character. For example, the **ResourceNames** property returns "Tamara,Tanya" if the list separator character is the comma (,) and the task has two resources named Tamara and Tanya. Project uses the list separator specified in the **Regional and Language Options** dialog box of the Windows Control Panel.


