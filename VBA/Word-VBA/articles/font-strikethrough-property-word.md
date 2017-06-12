---
title: Font.StrikeThrough Property (Word)
keywords: vbawd10.chm156369031
f1_keywords:
- vbawd10.chm156369031
ms.prod: word
api_name:
- Word.Font.StrikeThrough
ms.assetid: c55819cc-efb8-9981-3335-b3d6e6c26924
ms.date: 06/08/2017
---


# Font.StrikeThrough Property (Word)

 **True** if the font is formatted as strikethrough text. Read/write **Long** .


## Syntax

 _expression_ . **StrikeThrough**

 _expression_ An expression that returns a **[Font](font-object-word.md)** object.


## Remarks

The  **StrikeThrough** property returns **True** , **False** or **wdUndefined** (a mixture of **True** and **False** ). Can be set to **True** , **False** , or **wdToggle** .

To set or return double strikethrough formatting, use the  **[DoubleStrikeThrough](font-doublestrikethrough-property-word.md)** property.


## Example

This example applies strikethrough formatting to the first three words in the active document.


```vb
Set myDoc = ActiveDocument 
Set myRange = myDoc.Range(Start:=myDoc.Words(1).Start, _ 
 End:=myDoc.Words(3).End) 
myRange.Font.StrikeThrough = True
```

This example applies strikethrough formatting to the selected text.




```vb
If Selection.Type = wdSelectionNormal Then 
 Selection.Font.StrikeThrough = True 
Else 
 MsgBox "You need to select some text." 
End If
```


## See also


#### Concepts


[Font Object](font-object-word.md)

