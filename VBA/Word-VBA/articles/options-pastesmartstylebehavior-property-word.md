---
title: Options.PasteSmartStyleBehavior Property (Word)
keywords: vbawd10.chm162988464
f1_keywords:
- vbawd10.chm162988464
ms.prod: word
api_name:
- Word.Options.PasteSmartStyleBehavior
ms.assetid: 1d6723e1-7b25-87cd-7d08-622a0e734c2f
ms.date: 06/08/2017
---


# Options.PasteSmartStyleBehavior Property (Word)

 **True** if Microsoft Word intelligently merges styles when pasting a selection from a different document. Read/write **Boolean** .


## Syntax

 _expression_ . **PasteSmartStyleBehavior**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Example

This example sets Word to intelligently paste styles in text selected from a different document if the option has been disabled.


```vb
Sub UseSmartStyle() 
 With Options 
 If .PasteSmartStyleBehavior = False Then 
 .PasteSmartStyleBehavior = True 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

