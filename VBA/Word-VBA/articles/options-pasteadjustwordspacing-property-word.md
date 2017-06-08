---
title: Options.PasteAdjustWordSpacing Property (Word)
keywords: vbawd10.chm162988461
f1_keywords:
- vbawd10.chm162988461
ms.prod: word
api_name:
- Word.Options.PasteAdjustWordSpacing
ms.assetid: 28c20e9a-8ebe-323f-0fa5-63c6310e988e
ms.date: 06/08/2017
---


# Options.PasteAdjustWordSpacing Property (Word)

 **True** if Microsoft Word automatically adjusts the spacing of words when cutting and pasting selections. Read/write **Boolean** .


## Syntax

 _expression_ . **PasteAdjustWordSpacing**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Example

This example sets Word to automatically adjust the spacing of words when cutting and pasting selections if the option has been disabled.


```vb
Sub AdjustWordSpace() 
 With Options 
 If .PasteAdjustWordSpacing = False Then 
 .PasteAdjustWordSpacing = True 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

