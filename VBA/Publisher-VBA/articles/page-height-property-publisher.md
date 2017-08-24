---
title: Page.Height Property (Publisher)
keywords: vbapb10.chm393240
f1_keywords:
- vbapb10.chm393240
ms.prod: publisher
api_name:
- Publisher.Page.Height
ms.assetid: 7ab931d7-c4aa-4687-44f8-2d03a389cd4f
ms.date: 06/08/2017
---


# Page.Height Property (Publisher)

Returns a  **Long** that represent the height (in points) of a cell, range of cells, or page. Read-only.


## Syntax

 _expression_. **Height**

 _expression_A variable that represents a  **Page** object.


## Remarks

The valid range for the  **Height** property depends on the size of the application workspace and the position of the object within the workspace. For centered objects on non-banner page sizes, the **Height** property may be 0.0 to 50.0 inches. For centered objects on banner page sizes, the **Height** property may be 0.0 to 241.0 inches.


