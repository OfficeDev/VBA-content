---
title: Shape.CalloutTarget Property (Visio)
keywords: vis_sdr.chm11262470
f1_keywords:
- vis_sdr.chm11262470
ms.prod: visio
api_name:
- Visio.Shape.CalloutTarget
ms.assetid: 4366753a-c8e2-ba85-54fd-9c74cd21d762
ms.date: 06/08/2017
---


# Shape.CalloutTarget Property (Visio)

Gets or sets the target shape that is associated with the callout shape by a callout relationship. Read/write.


## Syntax

 _expression_ . **CalloutTarget**

 _expression_ A variable that represents a **[Shape](shape-object-visio.md)** object.


### Return Value

 **Shape**


## Remarks

If you attempt to get or set the  **CalloutTarget** property value on a shape that is not a callout, Microsoft Visio will return an **Inappropriate Source Object** error.

If no target shape is associated with the callout shape, the  **CalloutTarget** property returns **Nothing** . Setting the property value to **Nothing** removes any target shapes that are associated with the callout shape.


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **CalloutTarget** property to set the target shape of a callout.


```vb
Set vsoShape = vsoCalloutShape.CalloutTarget
```


