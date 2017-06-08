---
title: Options.EnableHangulHanjaRecentOrdering Property (Word)
keywords: vbawd10.chm162988374
f1_keywords:
- vbawd10.chm162988374
ms.prod: word
api_name:
- Word.Options.EnableHangulHanjaRecentOrdering
ms.assetid: 2b34789f-2bbb-b062-c3da-157f5d51cce8
ms.date: 06/08/2017
---


# Options.EnableHangulHanjaRecentOrdering Property (Word)

 **True** if Microsoft Word displays the most recently used words at the top of the suggestions list during conversion between Hangul and Hanja. Read/write **Boolean** .


## Syntax

 _expression_ . **EnableHangulHanjaRecentOrdering**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example asks the user whether to set Microsoft Word to display the most recently used words at the top of the suggestions list during conversion between Hangul and Hanja.


```vb
x = MsgBox("Display most recently used words " _ 
 &; "at the top of the suggestions list?", vbYesNo) 
If x = vbYes Then 
 Options.EnableHangulHanjaRecentOrdering = True 
Else 
 Options.EnableHangulHanjaRecentOrdering = False 
End If
```


## See also


#### Concepts


[Options Object](options-object-word.md)

