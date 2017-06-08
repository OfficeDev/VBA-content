---
title: MasterShortcut.ShapeHelp Property (Visio)
keywords: vis_sdr.chm16014325
f1_keywords:
- vis_sdr.chm16014325
ms.prod: visio
api_name:
- Visio.MasterShortcut.ShapeHelp
ms.assetid: 79a4c230-4f34-1644-6da3-bd72f116c11e
ms.date: 06/08/2017
---


# MasterShortcut.ShapeHelp Property (Visio)

Gets or sets the help string used when the user clicks  **Help** on the shortcut menu of a master shortcut. Read/write.


## Syntax

 _expression_ . **ShapeHelp**

 _expression_ A variable that represents a **MasterShortcut** object.


### Return Value

String


## Remarks

If the help string is blank, the  **Help** command uses the help string defined by the shortcut's target master, determined by the **Help** property of that master's top-level shape.


