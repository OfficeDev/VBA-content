---
title: Selection.ExtendMode Property (Word)
keywords: vbawd10.chm158663062
f1_keywords:
- vbawd10.chm158663062
ms.prod: word
api_name:
- Word.Selection.ExtendMode
ms.assetid: 7b12cf8b-9be1-6ebc-de96-e7734eaad3b6
ms.date: 06/08/2017
---


# Selection.ExtendMode Property (Word)

 **True** if Extend mode is active. Read/write **Boolean** .


## Syntax

 _expression_ . **ExtendMode**

 _expression_ An expression that returns a **[Selection](selection-object-word.md)** object.


## Remarks

When Extend mode is active, the Extend argument of the following methods is  **True** by default: **[EndKey](selection-endkey-method-word.md)** , **[HomeKey](selection-homekey-method-word.md)** , **[MoveDown](selection-movedown-method-word.md)** , **[MoveLeft](selection-moveleft-method-word.md)** , **[MoveRight](selection-moveright-method-word.md)** , and **[MoveUp](selection-moveup-method-word.md)** . Also, the letters "EXT" appear on the status bar.

This property can only be set during run time; attempts to set it in Immediate mode are ignored. The Extend arguments of the  **[EndOf](selection-endof-method-word.md)** and **[StartOf](selection-startof-method-word.md)** methods are not affected by this property.


## Example

This example moves to the beginning of the paragraph and selects the paragraph plus the next two sentences.


```vb
With Selection 
 .MoveUp Unit:=wdParagraph 
 .ExtendMode = True 
 .MoveDown Unit:=wdParagraph 
 .MoveRight Unit:=wdSentence, Count:=2 
End With
```

This example collapses the current selection, turns on Extend mode, and selects the current sentence.




```vb
With Selection 
 .Collapse 
 .ExtendMode = True 
 ' Select current word. 
 .Extend 
 ' Select current sentence. 
 .Extend 
End With
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

