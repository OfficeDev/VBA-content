---
title: Selection.DataGraphic Property (Visio)
keywords: vis_sdr.chm11160205
f1_keywords:
- vis_sdr.chm11160205
ms.prod: visio
api_name:
- Visio.Selection.DataGraphic
ms.assetid: 09275500-7b8a-2d78-971c-2e27bc3b9e46
ms.date: 06/08/2017
---


# Selection.DataGraphic Property (Visio)

Gets the data graphic master ( **Master** object of type **visTypeDataGraphic** ) of the primary shape in the selection. Sets the data graphic master of all shapes in the selection. Read/write.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **DataGraphic**

 _expression_ An expression that returns a **Selection** object.


### Return Value

Master


## Remarks

If the selection has no data graphic master associated with it, the  **DataGraphic** property returns **Nothing** .


