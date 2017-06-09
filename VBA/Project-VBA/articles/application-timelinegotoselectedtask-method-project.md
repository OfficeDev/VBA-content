---
title: Application.TimelineGotoSelectedTask Method (Project)
keywords: vbapj.chm61
f1_keywords:
- vbapj.chm61
ms.prod: project-server
api_name:
- Project.Application.TimelineGotoSelectedTask
ms.assetid: 62353aab-b850-bcf9-1d16-c7c794643318
ms.date: 06/08/2017
---


# Application.TimelineGotoSelectedTask Method (Project)

When a task is selected in the Timeline view,  **TimelineGotoSelectedTask** selects the same task in the main view.


## Syntax

 _expression_. **TimelineGotoSelectedTask**

 _expression_ An expression that returns an **Application** object.


### Return Value

 **Boolean**


## Remarks

The  **TimelineGotoSelectedTask** method corresponds to the **Go to Selected Task** command on the option menu in the Timeline view. After you run the **TimelineGotoSelectedTask** method, the Timeline view remains the active view.

If a single task is not selected in the Timeline view, or if the Timeline view is not active, the  **TimelineGotoSelectedTask** method results in run-time error 1100, "The method is not available in this situation."


