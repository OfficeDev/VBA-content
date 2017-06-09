---
title: Viewer.LayerColor Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.LayerColor
ms.assetid: 5e1bb40e-3e50-7ab9-a43d-606df8e0d14f
ms.date: 06/08/2017
---


# Viewer.LayerColor Property (Visio Viewer)

Gets or sets the color of the layer at the specified index position in the current drawing open in Microsoft Visio Viewer. Read/write.


## Syntax

 _expression_. **LayerColor**( **_LayerIndex_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|LayerIndex|Required| **Long**|The index of the layer in the collection of layers in the drawing open in Visio Viewer.|

### Return Value

 **[OLE_COLOR]**


## Remarks

Returns a value of data type  **OLE_COLOR** that represents the color of the specified layer in Visio Viewer. The **OLE_COLOR** data type is used for properties that return colors.

Valid hexadecimal values for an  **OLE_COLOR** data type in Visio Viewer are of the form _&;Hbbggrr_, where  _bb_ is the blue value, _gg_ the green value, and _rr_ the red value. All three color values range between 00 and FF hexadecimal (255 decimal).

The collection of layers is one-based, so the index of the first layer in the collection is 1. If there are no layers in the drawing, or if you pass the index of a nonexistent layer, the  **LayerColor** property returns 0.


## Example

The following code shows how to get the color of the layer at index position 1 in the drawing open in Visio Viewer.


```vb
Debug.Print vsoViewer.LayerColor(1)
```


