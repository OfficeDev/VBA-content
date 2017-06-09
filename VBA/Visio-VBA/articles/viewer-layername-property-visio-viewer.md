---
title: Viewer.LayerName Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.LayerName
ms.assetid: ebf2b8da-7c4d-b67c-9f8c-17629f1d8214
ms.date: 06/08/2017
---


# Viewer.LayerName Property (Visio Viewer)

Gets the name of the layer at the specified index in the drawing open in Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **LayerName**( **_LayerIndex_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|LayerIndex|Required| **Long**|The index of the layer in the collection of layers in the drawing open in Visio Viewer.|

### Return Value

 **String**


## Remarks

The collection of layers is one-based, so the index of the first layer in the collection is 1. If there are no layers in the drawing, or if there is no layer at index position LayerIndex, the  **LayerName** property returns nothing.


## Example

The following code gets the name of all the layers in the drawing open in Visio Viewer.


```vb
Dim intCounter As Integer

For intCounter = 1 To vsoViewer.LayerCount

    Debug.Print vsoViewer.LayerName(intCounter)

Next intCounter
```


