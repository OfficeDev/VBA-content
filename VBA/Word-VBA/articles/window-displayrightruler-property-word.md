---
title: Window.DisplayRightRuler Property (Word)
keywords: vbawd10.chm157417507
f1_keywords:
- vbawd10.chm157417507
ms.prod: word
api_name:
- Word.Window.DisplayRightRuler
ms.assetid: a587b652-5ba6-564d-4a8e-d78649bd716d
ms.date: 06/08/2017
---


# Window.DisplayRightRuler Property (Word)

 **True** if the vertical ruler appears on the right side of the document window in print layout view. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayRightRuler**

 _expression_ An expression that returns a **[Window](window-object-word.md)** object.


## Remarks

For more information on using Word with right-to-left languages, see Word features for right-to-left languages .


## Example

This example sets the active window to print layout view and displays the vertical ruler on the right side.


```vb
With ActiveWindow 
 .View = wdPrintView 
 .DisplayRightRuler = True 
End With
```


## See also


#### Concepts


[Window Object](window-object-word.md)

