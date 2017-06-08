---
title: Options.AutoCreateNewDrawings Property (Word)
keywords: vbawd10.chm162988483
f1_keywords:
- vbawd10.chm162988483
ms.prod: word
api_name:
- Word.Options.AutoCreateNewDrawings
ms.assetid: d774e700-d62d-1418-e860-b3cd05281468
ms.date: 06/08/2017
---


# Options.AutoCreateNewDrawings Property (Word)

 **True** for Microsoft Word to draw newly created shapes in a drawing canvas. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoCreateNewDrawings**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Remarks

The  **AutoCreateNewDrawings** property only affects shapes as they are added from within Word. If shapes are added through Visual Basic for Applications code, they are added as specified in the code regardless of whether this option is set to **True** or **False** .


## Example

This example sets Word to add newly created shapes directly to the document and not within a drawing canvas.


```vb
Sub NewDrawings() 
 Application.Options.AutoCreateNewDrawings = False 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

