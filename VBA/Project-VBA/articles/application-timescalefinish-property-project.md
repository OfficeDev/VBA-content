---
title: Application.TimescaleFinish Property (Project)
keywords: vbapj.chm132757
f1_keywords:
- vbapj.chm132757
ms.prod: project-server
api_name:
- Project.Application.TimescaleFinish
ms.assetid: 66c07ebc-ee68-bf4c-9af1-c894d4617e44
ms.date: 06/08/2017
---


# Application.TimescaleFinish Property (Project)

Gets the date and time that the timescale in the current view ends. Read-only  **Variant**.


## Syntax

 _expression_. **TimescaleFinish**

 _expression_ An expression that returns an **Application** object.


## Remarks

The end of the timescale in a Gantt chart can be moved to a position within the time period. To change the timescale duration, use any of the following methods:  **ZoomTimescale**,  **ZoomOut**,  **ZoomIn**, or  **Zoom**.


## Example

If the Gantt chart timescale ends on June 2, 2012, the following statement shows  **6/2/2012 2:51:00 PM** in the **Immediate** pane of the VBE.


```vb
Debug.Print TimescaleFinish
```


