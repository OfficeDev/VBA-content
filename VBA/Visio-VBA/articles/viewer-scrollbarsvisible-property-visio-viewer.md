---
title: Viewer.ScrollbarsVisible Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.ScrollbarsVisible
ms.assetid: cd8f5b2d-f604-8bac-2e82-338cfa7d7174
ms.date: 06/08/2017
---


# Viewer.ScrollbarsVisible Property (Visio Viewer)

Gets or sets a value that indicates whether the scroll bars are visible in Microsoft Visio Viewer. Read/write.


## Syntax

 _expression_. **ScrollbarsVisible**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **Boolean**


## Remarks

The default is for the scroll bars to be visible ( **True**).


## Example

The following code turns off display of the scroll bars in the drawing that is open in Visio Viewer.


```vb
vsoViewer.ScrollbarsVisible = False
```


