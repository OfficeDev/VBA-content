---
title: Viewer.Zoom Property (Visio Viewer)
ms.prod: visio
ms.assetid: 52bb7493-836e-1e1b-a91e-cb077f881c00
ms.date: 06/08/2017
---


# Viewer.Zoom Property (Visio Viewer)

Gets or sets the percentage of zoom for Microsoft Visio Viewer. Read/write.


## Syntax

 _expression_. **Zoom**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **Double**


## Remarks

Possible values for the  **Zoom** property range from 1% through 400%, and also include "Page", Width", and "Last".


## Example

The following code gets the percentage of zoom in the drawing that is open in Visio Viewer.


```vb
Debug.Print "Zoom = "; vsoViewer.Zoom
```


