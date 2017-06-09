---
title: Page.CanUndo Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 86494409-ae9f-4830-c7dd-f5e8284e04b0
ms.date: 06/08/2017
---


# Page.CanUndo Property (Outlook Forms Script)

Returns a  **Boolean** that specifies whether the last user action can be undone. Read-only.


## Syntax

 _expression_. **CanUndo**

 _expression_A variable that represents a  **Page** object.


## Remarks

 **True** if the most recent user action can be undone, **False** if the most recent user action cannot be undone.

 **CanUndo** is read-only.

Many user actions can be undone with the  **Undo** command. The **CanUndo** property indicates whether the most recent action can be undone.


