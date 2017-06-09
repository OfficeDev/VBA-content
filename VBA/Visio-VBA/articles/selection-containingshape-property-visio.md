---
title: Selection.ContainingShape Property (Visio)
keywords: vis_sdr.chm11113320
f1_keywords:
- vis_sdr.chm11113320
ms.prod: visio
api_name:
- Visio.Selection.ContainingShape
ms.assetid: c25dec03-dfa9-d61f-ad02-8ea7ee6cd87f
ms.date: 06/08/2017
---


# Selection.ContainingShape Property (Visio)

Returns the  **Shape** object that contains an object or collection. Read-only.


## Syntax

 _expression_ . **ContainingShape**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Shape


## Remarks

If the  **Shape** object is the member of a group, the **ContainingShape** property returns that group.

If the  **Shape** object is a top-level shape in its **Page** or **Master** object (it is not a member of a group), the **ContainingShape** property returns the page sheet of its page or master.

If the  **Shape** object is the page sheet of a page or master, the **ContainingShape** property returns **Nothing** .


