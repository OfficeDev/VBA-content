---
title: Shape.IsDataGraphicCallout Property (Visio)
ms.prod: visio
api_name:
- Visio.Shape.IsDataGraphicCallout
ms.assetid: dedf6880-e597-8582-12e5-18bfe6286e66
ms.date: 06/08/2017
---


# Shape.IsDataGraphicCallout Property (Visio)

Specifes whether a shape is a data graphic callout. Read-only.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **IsDataGraphicCallout**

 _expression_ An expression that returns a **Shape** object.


### Return Value

Boolean


## Remarks

If the parent shape is a data graphic callout, the  **IsDataGraphicCallout** method returns **True** . If the shape is not a data graphic callout, the method returns **False** . Note that the method also returns **False** if the parent shape is a sub-shape of a shape that is a data graphic callout.


