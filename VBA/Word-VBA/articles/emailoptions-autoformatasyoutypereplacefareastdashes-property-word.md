---
title: EmailOptions.AutoFormatAsYouTypeReplaceFarEastDashes Property (Word)
keywords: vbawd10.chm165347629
f1_keywords:
- vbawd10.chm165347629
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeReplaceFarEastDashes
ms.assetid: 0a2fbf7f-9f32-b1d9-03aa-7e43d3b8b8ec
ms.date: 06/08/2017
---


# EmailOptions.AutoFormatAsYouTypeReplaceFarEastDashes Property (Word)

 **True** for Microsoft Word to automatically correct long vowel sounds and dashes. Read/write.


## Syntax

 _expression_ . **AutoFormatAsYouTypeReplaceFarEastDashes**

 _expression_ Required. A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Example

This example sets Microsoft Word to automatically correct long vowel sounds and dashes as you type.


```vb
Sub AutoFarEastDashes() 
 Options.AutoFormatAsYouTypeReplaceFarEastDashes = True 
End Sub
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

