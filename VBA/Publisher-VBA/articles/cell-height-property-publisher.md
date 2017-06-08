---
title: Cell.Height Property (Publisher)
keywords: vbapb10.chm5111825
f1_keywords:
- vbapb10.chm5111825
ms.prod: publisher
api_name:
- Publisher.Cell.Height
ms.assetid: ced71ac0-eca8-0939-8812-fe0e79a47cba
ms.date: 06/08/2017
---


# Cell.Height Property (Publisher)

Returns a  **Long** that represent the height (in cells) of a table, range of cells, or page. Read-only.


## Syntax

 _expression_. **Height**

 _expression_A variable that represents a  **Cell** object.


## Remarks

The valid range for the  **Height** property depends on the size of the application workspace and the position of the object within the workspace. For centered objects on non-banner page sizes, the **Height** property may be 0.0 to 50.0 cells. For centered objects on banner page sizes, the **Height** property may be 0.0 to 241.0 cells.


