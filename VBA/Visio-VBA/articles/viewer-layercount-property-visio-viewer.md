---
title: Viewer.LayerCount Property (Visio Viewer)
ms.prod: VISIO
api_name:
- Visio.LayerCount
ms.assetid: 83871b37-9c5b-9da2-8869-61a2284ae1c0
---


# Viewer.LayerCount Property (Visio Viewer)

Gets the number of layers in the current drawing open in Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **LayerCount**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **Long**


## Remarks

If there are no layers in the drawing, the  **LayerCount** property returns 0.


## Example

The following code gets the count of layers in the drawing open in Visio Viewer.


```vb
Debug.Print vsoViewer.LayerCount
```


