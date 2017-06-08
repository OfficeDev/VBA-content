---
title: CommandBarEvents Property (VBA Add-In Object Model)
keywords: vbob6.chm100197
f1_keywords:
- vbob6.chm100197
ms.prod: office
ms.assetid: 342f5e9c-c5cc-4596-0b05-0985df1aba49
ms.date: 06/08/2017
---


# CommandBarEvents Property (VBA Add-In Object Model)



Returns the  **CommandBarEvents** object. Read-only.
 **Settings**
The setting for the [argument](vbe-glossary.md) you pass to the **CommandBarEvents** property is:


|**Argument**|**Description**|
|:-----|:-----|
| _vbcontrol_|Must be an object of type  **CommandBarControl**.|
 **Remarks**
Use the  **CommandBarEvents** property to return an[event source object](vbe-glossary.md) that triggers an event when a command bar button is clicked. The argument passed to the **CommandBarEvents** property is the command bar control for which the Click event will be triggered.

