---
title: Viewer.LayerVisible Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.LayerVisible
ms.assetid: b62ce57e-a1a0-06b2-ade5-71e1c11b0596
ms.date: 06/08/2017
---


# Viewer.LayerVisible Property (Visio Viewer)

Gets or sets a value that indicates whether the specified layer is visible in the drawing open in Microsoft Visio Viewer. Read/write.


## Syntax

 _expression_. **LayerVisible**( **_LayerIndex_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|LayerIndex|Required| **Long**|The index of the layer in the collection of layers in the drawing open in Visio Viewer.|

### Return Value

 **Boolean**


## Remarks

The collection of layers is one-based, so the index of the first layer in the collection is 1. If there are no layers in the drawing, the  **LayerVisible** property returns **False**.


## Example

The following code gets a value that indicates whether the layer at index position 1 is visible.


```vb
Debug.Print vsoViewer.LayerVisible(1)
```


