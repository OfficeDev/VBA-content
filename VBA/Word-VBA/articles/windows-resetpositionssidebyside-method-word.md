---
title: Windows.ResetPositionsSideBySide Method (Word)
keywords: vbawd10.chm157351950
f1_keywords:
- vbawd10.chm157351950
ms.prod: word
api_name:
- Word.Windows.ResetPositionsSideBySide
ms.assetid: f9741635-ecc5-77a1-51d6-d48ef42a3ce6
ms.date: 06/08/2017
---


# Windows.ResetPositionsSideBySide Method (Word)

Resets two document windows that are in the  **Compare side by side with** view mode.


## Syntax

 _expression_ . **ResetPositionsSideBySide**

 _expression_ Required. A variable that represents a **[Windows](windows-object-word.md)** collection.


## Remarks

This method corresponds to the  **Reset Window Position** button on the **Compare Side by Side** toolbar. Use the **ResetPositionsSideBySide** method to reset the display of two documents. For example, if a user minimizes or maximizes one of the two document windows being compared, the **ResetPositionsSideBySide** method resets the display so that the two windows are displayed side by side again.


## Example

The following example places two documents that were previously placed in side-by-side windows in adjacent windows.


```
Windows.ResetPositionsSideBySide
```


## See also


#### Concepts


[Windows Collection Object](windows-object-word.md)

