---
title: Application.SupportsMultipleWindows Property (Project)
keywords: vbapj.chm132628
f1_keywords:
- vbapj.chm132628
ms.prod: project-server
api_name:
- Project.Application.SupportsMultipleWindows
ms.assetid: d52eb74c-a809-2084-9e4e-45ca4d53d2e4
ms.date: 06/08/2017
---


# Application.SupportsMultipleWindows Property (Project)

Always  **True** for Project and any other application that can have more than one window open at a time. Read-only **Boolean**.


## Syntax

 _expression_. **SupportsMultipleWindows**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **SupportsMultipleWindows** property is useful with Automation. For example, suppose you want to open a second window in the application referred to by a variable. If the variable refers to one of several possible applications, you may want to use the **SupportsMultipleWindows** property to confirm that the application currently referenced by the variable can have more than one window open at a time.


