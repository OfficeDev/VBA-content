---
title: Range.Kana Property (Word)
keywords: vbawd10.chm157155655
f1_keywords:
- vbawd10.chm157155655
ms.prod: word
api_name:
- Word.Range.Kana
ms.assetid: ed64b73e-6970-3099-6f75-0beac6bba84e
ms.date: 06/08/2017
---


# Range.Kana Property (Word)

Returns or sets whether the specified range of Japanese language text is hiragana or katakana. Read/write  **WdKana** .


## Syntax

 _expression_ . **Kana**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

This property returns  **wdUndefined** if the range contains a mix of hiragana and katakana or if the range contains some non-Japanese text. If you set the **Kana** property to **wdUndefined** , an error occurs.


## Example

This example displays what type of Japanese text the current selection contains.


```vb
Select Case Selection.Range.Kana 
 Case wdKanaHiragana 
 MsgBox "This text is hiragana." 
 Case wdKanaKatakana 
 MsgBox "This text is katakana." 
 Case wdUndefined 
 MsgBox "This text is a mix of " _ 
 &; "hiragana and katakana." 
End Select
```


## See also


#### Concepts


[Range Object](range-object-word.md)

