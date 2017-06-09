---
title: Frame.CanUndo Property (Outlook Forms Script)
keywords: olfm10.chm2000870
f1_keywords:
- olfm10.chm2000870
ms.prod: outlook
ms.assetid: 7cb4090f-8886-17c9-2bd3-cdeb78e5aa57
ms.date: 06/08/2017
---


# Frame.CanUndo Property (Outlook Forms Script)

Returns a  **Boolean** that specifies whether the last user action can be undone. Read-only.


## Syntax

 _expression_. **CanUndo**

 _expression_A variable that represents a  **Frame** object.


## Remarks

 **True** if the most recent user action can be undone, **False** if the most recent user action cannot be undone.

 **CanUndo** is read-only.

Many user actions can be undone with the  **Undo** command. The **CanUndo** property indicates whether the most recent action can be undone.


