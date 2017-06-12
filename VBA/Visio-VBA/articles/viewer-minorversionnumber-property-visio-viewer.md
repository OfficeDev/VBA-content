---
title: Viewer.MinorVersionNumber Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.MinorVersionNumber
ms.assetid: 9bb21829-690c-0ff3-1e30-f9964cc903c4
ms.date: 06/08/2017
---


# Viewer.MinorVersionNumber Property (Visio Viewer)

Gets the minor version number of Microsoft Visio Viewer. Read-only.


## Syntax

 _expression_. **MinorVersionNumber**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **Long**


## Remarks

The minor version number represents the digit or digits that follow the major version number, separated by a period.


## Example

The following code displays the minor version number of Visio Viewer in the  **Immediate** window.


```vb
Debug.Print vsoViewer.MinorVersionNumber
```


