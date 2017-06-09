---
title: Page.LayoutRoutePassive Property (Visio)
keywords: vis_sdr.chm10962445
f1_keywords:
- vis_sdr.chm10962445
ms.prod: visio
api_name:
- Visio.Page.LayoutRoutePassive
ms.assetid: 7244abb5-0c8f-d68b-4b2d-3e192afe1d80
ms.date: 06/08/2017
---


# Page.LayoutRoutePassive Property (Visio)

Determines whether to enable advanced connector routing logic on the page. Read/write.


## Syntax

 _expression_ . **LayoutRoutePassive**

 _expression_ A variable that represents a **[Page](page-object-visio.md)** object.


### Return Value

 **Boolean**


## Remarks

When  **LayoutRoutePassive** is set to **True** , advanced connector routing is disabled, which means that when connectors are glued to shapes, they do not attach to the shapes dynamically. That is, when those shapes are moved, the connector endpoints that are glued to the shapes move with them, but without the benefits of dynamic routing, such as shifting their location to a new side of the shape when appropriate, or redistribution (in the Flowchart routing style only).

The default setting of the  **LayoutRoutePassive** property is **False** , which means that advanced connector routing is enabled. Because setting the property to **True** disables advanced connector routing, doing so can result in shapes that are glued incorrectly, which, in turn, leads to unsatisfactory diagram appearance and connector behavior. As a result, it is suggested that you set this property to **True** only briefly, when you specifically do not want advanced connector routing to occur.

The  **LayoutRoutePassive** property can also affect paste behavior. When you paste shapes into a diagram, the resulting connector behavior depends on how you pasted the shape into the diagram. When you paste by using the keyboard combination CTRL+V, advanced connector rerouting does not take place, but when you paste by using the right-click (context) menu, advanced connector rerouting does take place. However, if the **LayoutRoutePassive** property is set to **True** , advanced connector routing is always disabled when you paste shapes into the drawing, no matter which method you use.


