---
title: Options.AutoFormatAsYouTypeDeleteAutoSpaces Property (Word)
keywords: vbawd10.chm162988334
f1_keywords:
- vbawd10.chm162988334
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeDeleteAutoSpaces
ms.assetid: a0308511-e676-73d5-cbe9-41ed3858828a
ms.date: 06/08/2017
---


# Options.AutoFormatAsYouTypeDeleteAutoSpaces Property (Word)

 **True** for Microsoft Word to automatically delete spaces inserted between Japanese and Latin text as you type. Read/write.


## Syntax

 _expression_ . **AutoFormatAsYouTypeDeleteAutoSpaces**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets Microsoft Word to automatically delete spaces inserted between Japanese and Latin text as you type.


```vb
Sub AutoDeleteSpaces() 
 Options.AutoFormatAsYouTypeDeleteAutoSpaces = True 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

