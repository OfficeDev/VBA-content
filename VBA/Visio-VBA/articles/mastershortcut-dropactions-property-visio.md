---
title: MasterShortcut.DropActions Property (Visio)
keywords: vis_sdr.chm16013450
f1_keywords:
- vis_sdr.chm16013450
ms.prod: visio
api_name:
- Visio.MasterShortcut.DropActions
ms.assetid: 6c835662-0ae4-4058-6fb9-7299f898150a
ms.date: 06/08/2017
---


# MasterShortcut.DropActions Property (Visio)

Defines special actions to be performed on shapes created by using a master shortcut. Read/write.


## Syntax

 _expression_ . **DropActions**

 _expression_ A variable that represents a **MasterShortcut** object.


### Return Value

String


## Remarks

When you drag a master shortcut onto a drawing page, Microsoft Visio applies any drop actions in the shortcut to the newly created shape. Each drop action defines a particular value or formula to be assigned to a particular cell in the new shape.

Because drop actions are defined by the shortcut, not the target master, it is possible to create several shortcuts that refer to the same target master, but which produce very different effects when dropped onto the drawing page.

The  **DropActions** property can be blank, or it can define a series of one or more individual drop actions. Actions are separated by semicolons (;). Each action consists of the name of the cell to change, followed by the formula to apply to that cell, separated by an equals sign (=). For example:




```
Angle=45; FillForegnd=7; Width=(ThePage!PageWidth / 2 - 4cm)
```

The application does not validate drop actions until they are applied to a new shape. If the  **DropActions** property contains syntax errors or invalid cell names, the offending actions are ignored. However, if the application is running in developer mode, an error message is displayed that identifies the invalid action and the cause of the error. When using shortcut drop actions in your code, always test your shortcuts in[developer mode](http://msdn.microsoft.com/library/1bdc55f5-8fc7-7257-03d5-c049eceb29ff%28Office.15%29.aspx) to make sure the drop actions do not contain errors.


