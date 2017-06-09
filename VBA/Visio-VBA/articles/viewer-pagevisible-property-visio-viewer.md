---
title: Viewer.PageVisible Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.PageVisible
ms.assetid: 7af34d35-b83d-931a-7116-fef8dab42f22
ms.date: 06/08/2017
---


# Viewer.PageVisible Property (Visio Viewer)

Gets or sets a value that indicates whether the drawing page is visible in Microsoft Visio Viewer. Read/write.


## Syntax

 _expression_. **PageVisible**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **Boolean**


## Remarks

The default is that the drawing page not be visible ( **False**).


## Example

The following example shows how to make the drawing page visible in Visio Viewer.


```vb
vsoViewer.PageVisible = True
```


