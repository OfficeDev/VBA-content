---
title: Frame.RedoAction Method (Outlook Forms Script)
keywords: olfm10.chm2000340
f1_keywords:
- olfm10.chm2000340
ms.prod: outlook
ms.assetid: d681d6e8-935b-f5f0-aaba-e5f63e7491bb
ms.date: 06/08/2017
---


# Frame.RedoAction Method (Outlook Forms Script)

Reverses the effect of the most recent  **Undo** action.


## Syntax

 _expression_. **RedoAction**

 _expression_A variable that represents a  **Frame** object.


### Return Value

A  **Boolean** that specifies **True** if the method succeeds, **False** otherwise.


## Remarks

Redo reverses the last  **Undo**, which is not necessarily the last action taken. Not all actions can be undone.


