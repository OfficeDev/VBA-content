---
title: Font.Superscript Property (Word)
keywords: vbawd10.chm156369035
f1_keywords:
- vbawd10.chm156369035
ms.prod: word
api_name:
- Word.Font.Superscript
ms.assetid: 07a1f270-263e-00af-ed8f-3b9d2904e78b
ms.date: 06/08/2017
---


# Font.Superscript Property (Word)

 **True** if the font is formatted as superscript. Read/write **Long** .


## Syntax

 _expression_ . **Superscript**

 _expression_ An expression that returns a **[Font](font-object-word.md)** object.


## Remarks

Returns  **True** , **False** , or **wdUndefined** (a mixture of **True** and **False** ). Can be set to **True** , **False** , or **wdToggle** .

Setting the  **Superscript** property to **True** sets the **[Subscript](font-subscript-property-word.md)** property to **False** , and vice versa.


## Example

This example inserts text at the beginning of the active document and formats two characters in the fourth word as superscript.


```vb
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
myRange.InsertAfter "Superscript in the 4th word." 
ActiveDocument.Range(Start:=20, End:=22).Font.Superscript = True
```

This example formats the selected text as superscript.




```vb
If Selection.Type = wdSelectionNormal Then 
 Selection.Font.Superscript = True 
Else 
 MsgBox "You need to select some text." 
End If
```


## See also


#### Concepts


[Font Object](font-object-word.md)

