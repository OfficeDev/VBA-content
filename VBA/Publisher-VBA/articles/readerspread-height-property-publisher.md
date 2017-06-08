---
title: ReaderSpread.Height Property (Publisher)
keywords: vbapb10.chm524296
f1_keywords:
- vbapb10.chm524296
ms.prod: publisher
api_name:
- Publisher.ReaderSpread.Height
ms.assetid: dfb84798-da3f-516b-22cd-0ba2a63ff39d
ms.date: 06/08/2017
---


# ReaderSpread.Height Property (Publisher)

Returns a  **Single** that represents the height, in points, of the page. Read-only.


## Syntax

 _expression_. **Height**

 _expression_A variable that represents a  **ReaderSpread** object.


## Remarks

The valid range for the  **Height** property depends on the size of the application workspace and the position of the object within the workspace. For centered objects on non-banner page sizes, the **Height** property may be 0.0 to 50.0 inches. For centered objects on banner page sizes, the **Height** property may be 0.0 to 241.0 inches.


