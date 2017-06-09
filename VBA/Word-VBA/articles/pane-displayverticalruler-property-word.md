---
title: Pane.DisplayVerticalRuler Property (Word)
keywords: vbawd10.chm157286405
f1_keywords:
- vbawd10.chm157286405
ms.prod: word
api_name:
- Word.Pane.DisplayVerticalRuler
ms.assetid: 66899d6f-8e78-6d54-e0b0-d4a2bace428e
ms.date: 06/08/2017
---


# Pane.DisplayVerticalRuler Property (Word)

 **True** if a vertical ruler is displayed for the specified pane. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayVerticalRuler**

 _expression_ A variable that represents a **[Pane](pane-object-word.md)** object.


## Remarks

A vertical ruler appears only in print layout view, and only if the  **DisplayRulers** property is set to **True** .


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

