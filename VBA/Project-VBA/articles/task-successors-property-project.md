---
title: Task.Successors Property (Project)
keywords: vbapj.chm132674
f1_keywords:
- vbapj.chm132674
ms.prod: project-server
api_name:
- Project.Task.Successors
ms.assetid: 7e294395-00a7-ca80-ef58-506fbba1c9a8
ms.date: 06/08/2017
---


# Task.Successors Property (Project)

Gets or sets a list of the identification numbers of a task's successors. Read/write  **String**.


## Syntax

 _expression_. **Successors**

 _expression_ A variable that represents a **Task** object.


## Remarks

If the successors of the specified task have identification numbers of 2 and 10, and the list separator character is the comma, the  **Successors** property returns "2,10"


