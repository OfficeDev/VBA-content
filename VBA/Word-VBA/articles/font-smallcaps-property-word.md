---
title: Font.SmallCaps Property (Word)
keywords: vbawd10.chm156369029
f1_keywords:
- vbawd10.chm156369029
ms.prod: word
api_name:
- Word.Font.SmallCaps
ms.assetid: f2b0c4c9-2270-cb60-6bb1-fe2f4c183550
ms.date: 06/08/2017
---


# Font.SmallCaps Property (Word)

 **True** if the font is formatted as small capital letters. Read/write **Long** .


## Syntax

 _expression_ . **SmallCaps**

 _expression_ An expression that returns a **[Font](font-object-word.md)** object.


## Remarks

Returns  **True** , **False** or **wdUndefined** (a mixture of **True** and **False** ). Can be set to **True** , **False** , or **wdToggle** .

Setting the  **SmallCaps** property to **True** sets the **[AllCaps](font-allcaps-property-word.md)** property to **False** , and vice versa.


## Example

This example demonstrates the difference between small capital letters and all capital letters in a new document.


```vb
Set myRange = Documents.Add.Content 
With myRange 
 .InsertAfter "This is a demonstration of SmallCaps." 
 .Words(6).Font.SmallCaps = True 
 .InsertParagraphAfter 
 .InsertAfter "This is a demonstration of AllCaps." 
 .Words(14).Font.AllCaps = True 
End With
```

This example formats the entire selection as small capital letters if part of the selection is already formatted as small capital letters.




```vb
If Selection.Type = wdSelectionNormal Then 
 mySel = Selection.Font.SmallCaps 
 If mySel = wdUndefined Then Selection.Font.SmallCaps = True 
Else 
 MsgBox "You need to select some text." 
End If
```


## See also


#### Concepts


[Font Object](font-object-word.md)

