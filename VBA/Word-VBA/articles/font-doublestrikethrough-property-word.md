---
title: Font.DoubleStrikeThrough Property (Word)
keywords: vbawd10.chm156369032
f1_keywords:
- vbawd10.chm156369032
ms.prod: word
api_name:
- Word.Font.DoubleStrikeThrough
ms.assetid: 153d23c7-d5ee-4004-c540-ff23e263d9c5
ms.date: 06/08/2017
---


# Font.DoubleStrikeThrough Property (Word)

 **True** if the specified font is formatted as double strikethrough text. .


## Syntax

 _expression_ . **DoubleStrikeThrough**

 _expression_ A variable that represents a **[Font](font-object-word.md)** object.


## Remarks

Returns  **True** , **False** , or **wdUndefined** (a mixture of **True** and **False** ). Can be set to **True** , **False** , or **wdToggle** . Read/write **Long** .To set or return single-line strikethrough formatting, use the **[StrikeThrough](font-strikethrough-property-word.md)** property. Setting **DoubleStrikeThrough** to **True** sets **StrikeThrough** to **False** , and vice versa.


## Example

This example applies double strikethrough formatting to the selected text.


```vb
If Selection.Type = wdSelectionNormal Then 
 Selection.Font.DoubleStrikeThrough = True 
Else 
 MsgBox "You need to select some text." 
End If
```

This example removes double strikethrough formatting from the first word in the active document and capitalizes the first letter in the word.




```vb
With ActiveDocument.Words(1) 
 .Font.DoubleStrikeThrough = False 
 .Case = wdTitleSentence 
End With
```


## See also


#### Concepts


[Font Object](font-object-word.md)

