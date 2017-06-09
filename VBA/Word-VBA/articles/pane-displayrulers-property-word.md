---
title: Pane.DisplayRulers Property (Word)
keywords: vbawd10.chm157286404
f1_keywords:
- vbawd10.chm157286404
ms.prod: word
api_name:
- Word.Pane.DisplayRulers
ms.assetid: 25f30f02-41ff-1290-e10d-4b03df35e13f
ms.date: 06/08/2017
---


# Pane.DisplayRulers Property (Word)

 **True** if rulers are displayed for the specified pane. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayRulers**

 _expression_ A variable that represents a **[Pane](pane-object-word.md)** object.


## Remarks

The  **DisplayRulers** property is equivalent to the **Ruler** command on the **View** menu. If **DisplayRulers** is **False** , the horizontal and vertical rulers won't be displayed, regardless of the state of the **DisplayVerticalRuler** property.


## Example

This example switches the active pane to print layout view and displays the horizontal and vertical rulers.


```vb
With ActiveDocument.ActiveWindow.ActivePane 
 .View.Type = wdPrintView 
 .DisplayRulers = True 
 .DisplayVerticalRuler = True 
End With
```


## See also


#### Concepts


[Pane Object](pane-object-word.md)

