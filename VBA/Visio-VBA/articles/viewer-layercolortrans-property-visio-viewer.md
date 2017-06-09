---
title: Viewer.LayerColorTrans Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.LayerColorTrans
ms.assetid: b081bc59-701f-3ac5-6324-9eb055185c09
ms.date: 06/08/2017
---


# Viewer.LayerColorTrans Property (Visio Viewer)

Gets or sets the transparency of the color of the layer at the specified index in the current drawing in Microsoft Visio Viewer. Read/write.


## Syntax

 _expression_. **LayerColorTrans**( **_LayerIndex_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|LayerIndex|Required| **Long**|The index of the layer in the collection of layers in the drawing open in Visio Viewer.|

### Return Value

 **Double**


## Remarks

The collection of layers is one-based, so the index of the first layer in the collection is 1. If there are no layers in the drawing, or if you pass the index of a nonexistent layer, the  **LayerColorTrans** property returns 0.

The range of values for the  **LayerColorTrans** property is 0 to 1, which corresponds to a transparency percentage range of 0% to 100%. The property applies only to the active page in Visio Viewer.


## Example

The following code gets the transparency of the layer at index position 1.


```vb
Debug.Print vsoViewer.LayerColorTrans(1)
```


