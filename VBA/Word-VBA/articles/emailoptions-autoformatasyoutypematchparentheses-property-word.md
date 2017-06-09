---
title: EmailOptions.AutoFormatAsYouTypeMatchParentheses Property (Word)
keywords: vbawd10.chm165347628
f1_keywords:
- vbawd10.chm165347628
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeMatchParentheses
ms.assetid: bdb04e6e-a979-341c-fe6c-f7de33c1b568
ms.date: 06/08/2017
---


# EmailOptions.AutoFormatAsYouTypeMatchParentheses Property (Word)

 **True** for Microsoft Word to automatically correct improperly paired parentheses. Read/write.


## Syntax

 _expression_ . **AutoFormatAsYouTypeMatchParentheses**

 _expression_ Required. A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Example

This example sets Microsoft Word to automatically correct improperly paired parentheses as you type.


```vb
Sub AutoMatchParentheses() 
 Options.AutoFormatAsYouTypeMatchParentheses = True 
End Sub
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

