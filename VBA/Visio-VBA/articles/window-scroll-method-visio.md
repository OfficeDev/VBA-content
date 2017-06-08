---
title: Window.Scroll Method (Visio)
keywords: vis_sdr.chm11616520
f1_keywords:
- vis_sdr.chm11616520
ms.prod: visio
api_name:
- Visio.Window.Scroll
ms.assetid: 7d30ce8f-03b1-26ff-1495-efb9213083fa
ms.date: 06/08/2017
---


# Window.Scroll Method (Visio)

Scrolls the contents of a window vertically, horizontally, or both.


## Syntax

 _expression_ . **Scroll**( **_nxFlags_** , **_nyFlags_** )

 _expression_ A variable that represents a **Window** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _nxFlags_|Required| **Long**|Indicates how to scroll horizontally.|
| _nyFlags_|Required| **Long**|Indicates how to scroll vertically.|

### Return Value

Nothing


## Remarks

If the  **Window** object is not a built-in MDI or built-in docked stencil type, the **Scroll** method raises an exception.

Constants representing ways to scroll horizontally are declared by the Visio type library in  **[VisWindowScrollX](viswindowscrollx-enumeration-visio.md)** , and constants representing ways to scroll vertically are declared in **[VisWindowScrollY](viswindowscrolly-enumeration-visio.md)** .


