---
title: Viewer.PageCount Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.PageCount
ms.assetid: 3a7f90c0-6573-7ba5-414d-ede5b9c2fac6
ms.date: 06/08/2017
---


# Viewer.PageCount Property (Visio Viewer)

Gets the number of pages in the current document that is open in Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **PageCount**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **Long**


## Example

The following code displays the number of pages in the current document in the  **Immediate** window.


```vb
Debug.Print vsoViewer.PageCount
```


