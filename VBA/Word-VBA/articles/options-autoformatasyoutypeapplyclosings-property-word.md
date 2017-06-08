---
title: Options.AutoFormatAsYouTypeApplyClosings Property (Word)
keywords: vbawd10.chm162988331
f1_keywords:
- vbawd10.chm162988331
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeApplyClosings
ms.assetid: 179decd2-69b0-c734-3257-7d212894a5d2
ms.date: 06/08/2017
---


# Options.AutoFormatAsYouTypeApplyClosings Property (Word)

 **True** for Microsoft Word to automatically apply the Closing style to letter closings as you type. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeApplyClosings**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets Microsoft Word to automatically apply the Closing style to letter closings as you type.


```vb
Sub AutoClosings() 
 Options.AutoFormatAsYouTypeApplyClosings = True 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

