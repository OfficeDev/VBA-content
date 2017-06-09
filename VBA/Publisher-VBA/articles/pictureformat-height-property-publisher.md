---
title: PictureFormat.Height Property (Publisher)
keywords: vbapb10.chm3604759
f1_keywords:
- vbapb10.chm3604759
ms.prod: publisher
api_name:
- Publisher.PictureFormat.Height
ms.assetid: d98c76cc-4b75-28b7-5be1-101b372472d5
ms.date: 06/08/2017
---


# PictureFormat.Height Property (Publisher)

Returns a  **Variant** that represents the height, in points, of the specified picture or OLE object. Read-only.


## Syntax

 _expression_. **Height**

 _expression_A variable that represents a  **PictureFormat** object.


## Remarks

The valid range for the  **Height** property depends on the size of the application workspace and the position of the object within the workspace. For centered objects on non-banner page sizes, the **Height** property may be 0.0 to 50.0 inches. For centered objects on banner page sizes, the **Height** property may be 0.0 to 241.0 inches.


