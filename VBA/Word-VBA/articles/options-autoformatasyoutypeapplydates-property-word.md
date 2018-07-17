---
title: Options.AutoFormatAsYouTypeApplyDates Property (Word)
keywords: vbawd10.chm162988330
f1_keywords:
- vbawd10.chm162988330
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeApplyDates
ms.assetid: b31f13fa-9a76-3a86-c4c2-4720fec1b66b
ms.date: 06/08/2017
---


# Options.AutoFormatAsYouTypeApplyDates Property (Word)

 **True** for Microsoft Word to automatically apply the Date style to dates as you type. Read/write.


## Syntax

 _expression_ . **AutoFormatAsYouTypeApplyDates**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets Microsoft Word to automatically apply the Date style to dates as you type.


```vb
Sub AutoApplyDates() 
 Options.AutoFormatAsYouTypeApplyDates = True 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

