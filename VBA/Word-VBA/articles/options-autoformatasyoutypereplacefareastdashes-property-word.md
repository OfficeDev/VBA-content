---
title: Options.AutoFormatAsYouTypeReplaceFarEastDashes Property (Word)
keywords: vbawd10.chm162988333
f1_keywords:
- vbawd10.chm162988333
ms.prod: word
api_name:
- Word.Options.AutoFormatAsYouTypeReplaceFarEastDashes
ms.assetid: 2126603f-5cc6-4cb7-7a4e-1aac6b22571f
ms.date: 06/08/2017
---


# Options.AutoFormatAsYouTypeReplaceFarEastDashes Property (Word)

 **True** for Microsoft Word to automatically correct long vowel sounds and dashes. Read/write.


## Syntax

 _expression_ . **AutoFormatAsYouTypeReplaceFarEastDashes**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets Microsoft Word to automatically correct long vowel sounds and dashes as you type.


```vb
Sub AutoFarEastDashes() 
 Options.AutoFormatAsYouTypeReplaceFarEastDashes = True 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

