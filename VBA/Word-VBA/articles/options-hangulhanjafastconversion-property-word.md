---
title: Options.HangulHanjaFastConversion Property (Word)
keywords: vbawd10.chm162988372
f1_keywords:
- vbawd10.chm162988372
ms.prod: word
api_name:
- Word.Options.HangulHanjaFastConversion
ms.assetid: 3816fb7e-61e9-e2d8-bb03-c904130b1f10
ms.date: 06/08/2017
---


# Options.HangulHanjaFastConversion Property (Word)

 **True** if Microsoft Word automatically converts a word with only one suggestion during conversion between Hangul and Hanja. Read/write **Boolean** .


## Syntax

 _expression_ . **HangulHanjaFastConversion**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example asks the user whether to set Microsoft Word to use fast conversion during conversion between Hangul and Hanja.


```vb
x = MsgBox("Use fast conversion?", vbYesNo) 
If x = vbYes Then 
 Options.HangulHanjaFastConversion = True 
Else 
 Options.HangulHanjaFastConversion = False 
End If
```


## See also


#### Concepts


[Options Object](options-object-word.md)

