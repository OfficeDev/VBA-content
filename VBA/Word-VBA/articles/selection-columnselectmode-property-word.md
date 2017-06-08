---
title: Selection.ColumnSelectMode Property (Word)
keywords: vbawd10.chm158663063
f1_keywords:
- vbawd10.chm158663063
ms.prod: word
api_name:
- Word.Selection.ColumnSelectMode
ms.assetid: de146d32-63aa-3a17-6eeb-32cccf3f8bfd
ms.date: 06/08/2017
---


# Selection.ColumnSelectMode Property (Word)

 **True** if column selection mode is active. Read/write **Boolean** .


## Syntax

 _expression_ . **ColumnSelectMode**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

When this mode is active, the letters "COL" appear on the status bar.


## Example

This example selects a column of text that's two words across and three lines deep. The example copies the selection to the Clipboard and cancels column selection mode.


```vb
With Selection 
 .Collapse Direction:=wdCollapseStart 
 .ColumnSelectMode = True 
 .MoveRight Unit:=wdWord, Count:=2, Extend:=wdExtend 
 .MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend 
 .Copy 
 .ColumnSelectMode = False 
End With
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

