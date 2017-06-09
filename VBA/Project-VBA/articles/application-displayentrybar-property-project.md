---
title: Application.DisplayEntryBar Property (Project)
keywords: vbapj.chm131730
f1_keywords:
- vbapj.chm131730
ms.prod: project-server
api_name:
- Project.Application.DisplayEntryBar
ms.assetid: 56121152-2302-9d32-3a64-68b8b68f0f90
ms.date: 06/08/2017
---


# Application.DisplayEntryBar Property (Project)

Gets or sets a value that determines whether the data entry bar is visible.  **True** if the data entry bar is visible. Read/write **Boolean**.


## Syntax

 _expression_. **DisplayEntryBar**

 _expression_ A variable that represents an **Application** object.


## Remarks

If the entry bar is selected and you run the command  `DisplayEntryBar = False`, Project shows run-time error 1100, "The method is not available in this situation."

The  **DisplayEntryBar** property corresponds to the **Entry bar** checkbox on the **Display** tab of the **Project Options** dialog box.


