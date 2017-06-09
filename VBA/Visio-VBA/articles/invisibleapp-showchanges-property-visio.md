---
title: InvisibleApp.ShowChanges Property (Visio)
keywords: vis_sdr.chm17514690
f1_keywords:
- vis_sdr.chm17514690
ms.prod: visio
api_name:
- Visio.InvisibleApp.ShowChanges
ms.assetid: 8d38e592-366f-445f-9882-7b5864b5fa9e
ms.date: 06/08/2017
---


# InvisibleApp.ShowChanges Property (Visio)

Determines whether the screen is updated (redrawn) during a series of actions. Read/write.


## Syntax

 _expression_ . **ShowChanges**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Boolean


## Remarks

Use the  **ShowChanges** property to increase performance during a series of actions. For example, you can set the **ShowChanges** property to **False** while a series of shapes are created so that the screen is not redrawn after each shape appears. Then you can set it to **True** to update the screen.

If a program neglects to turn the  **ShowChanges** property on after turning it off, the Microsoft Visio instance will turn it back on when the user performs an operation.

The  **ShowChanges** property is similar to the **ScreenUpdating** property, which was implemented in Visio 3.0. In most cases using the **ShowChanges** property is preferable to using the **ScreenUpdating** property. Setting the **ShowChanges** property automatically sets the **ScreenUpdating** property; however, setting the **ScreenUpdating** property does not set the **ShowChanges** property.




- When  **ShowChanges** is **False** , the Visio instance will not refresh the screen (repaint drawing windows) as documents change or when they become obscured by other windows. All shapes in drawing and stencil windows are deselected and the Visio instance will not allow programs to change the selections of windows.
    
- When only  **ScreenUpdating** is **False** , the Visio instance will occasionally refresh the screen as documents change. **ScreenUpdating** does not cause deselects to occur or restrict selection changes.
    


The Visio instance will usually run faster when both the  **ShowChanges** and **ScreenUpdating** properties are **False** than when only the **ScreenUpdating** property is **False** . When both the **ShowChanges** and **ScreenUpdating** properties are **False** , the Visio views will not react to document changes until the **ShowChanges** property becomes **True** . This can cause noticeable delays after a program has completed a sequence of many operations. To cause some changes to occur as they happen, set **ScreenUpdating** to **True** immediately after setting **ShowChanges** to **False** . This can shorten the delay that occurs after **ShowChanges** becomes **True** , but will probably lengthen the time to complete the overall sequence of actions.


