---
title: DrawingControl.ShutdownBehavior Property (Visio)
ms.prod: VISIO
ms.assetid: 19c3e160-4b1d-40f1-b41d-69f21fca1d0d
---


# DrawingControl.ShutdownBehavior Property (Visio)

Determines how the Visio Drawing Control unloads the Visio application when the  **DrawingControl** object is released. Read/write **Integer**.


## Syntax

 _expression_ . **ShutdownBehavior**

 _expression_ A variable that represents a **DrawingControl** object.


### Return value

 **Integer**


## Remarks

A value of 0 (the default) does not unload MSO dlls when the drawing control is released. A value of 1 unloads the Visio application and MSO dlls when the drawing control is released.


## See also


#### Concepts


[DrawingControl Object](drawingcontrol-object-visio.md)

