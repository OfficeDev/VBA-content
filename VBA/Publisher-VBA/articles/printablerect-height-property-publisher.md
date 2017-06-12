---
title: PrintableRect.Height Property (Publisher)
keywords: vbapb10.chm7536646
f1_keywords:
- vbapb10.chm7536646
ms.prod: publisher
api_name:
- Publisher.PrintableRect.Height
ms.assetid: 55d07c00-ee9f-c177-3277-9355618dce6d
ms.date: 06/08/2017
---


# PrintableRect.Height Property (Publisher)

Returns a  **Single** that represents the height, in points, of the printable rectangle. Read-only.


## Syntax

 _expression_. **Height**

 _expression_A variable that represents a  **PrintableRect** object.


## Remarks

The valid range for the  **Height** property depends on the size of the application workspace and the position of the object within the workspace. For centered objects on non-banner page sizes, the **Height** property may be 0.0 to 50.0 inches. For centered objects on banner page sizes, the **Height** property may be 0.0 to 241.0 inches.


