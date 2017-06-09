---
title: Viewer.GridVisible Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.GridVisible
ms.assetid: 77351c96-c796-5a58-51ed-552843172ec0
ms.date: 06/08/2017
---


# Viewer.GridVisible Property (Visio Viewer)

Gets or sets a value that indicates whether the page grid is visible in Microsoft Visio Viewer. Read/write.


## Syntax

 _expression_. **GridVisible**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **Boolean**


## Remarks

You can display the grid only when the page is visible (that is, when the  **PageVisible** property is set to **True**, its default setting).


## Example

The following code shows how to display the grid in Visio Viewer.


```vb
vsoViewer.GridVisible = True
```


