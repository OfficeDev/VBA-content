---
title: Viewer.ToolbarVisible Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.ToolbarVisible
ms.assetid: 55e6b5fc-bda6-fff4-9049-b4aa398a4744
ms.date: 06/08/2017
---


# Viewer.ToolbarVisible Property (Visio Viewer)

Gets or sets a value that indicates whether the toolbar is visible in Microsoft Visio Viewer. Read/write.


## Syntax

 _expression_. **ToolbarVisible**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **Boolean**


## Remarks

The default is for the toolbar to be visible ( **True**).


## Example

The following code hides the toolbar in Visio Viewer.


```vb
vsoViewer.ToolbarVisible = False
```


