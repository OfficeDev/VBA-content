---
title: Font.Subscript Property (Word)
keywords: vbawd10.chm156369034
f1_keywords:
- vbawd10.chm156369034
ms.prod: word
api_name:
- Word.Font.Subscript
ms.assetid: 51226088-218d-4848-1358-d524fb2fe56a
ms.date: 06/08/2017
---


# Font.Subscript Property (Word)

 **True** if the font is formatted as subscript. Read/write **Long** .


## Syntax

 _expression_ . **Subscript**

 _expression_ An expression that returns a **[Font](font-object-word.md)** object.


## Remarks

Returns  **True** , **False** or **wdUndefined** (a mixture of **True** and **False** ). Can be set to **True** , **False** , or **wdToggle** .

Setting the  **Subscript** property to **True** sets the **[Superscript](font-superscript-property-word.md)** property to **False** , and vice versa.


## Example

This example inserts text at the beginning of the active document and formats the tenth character as subscript.


```vb
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
myRange.InsertAfter "Water = H20" 
myRange.Characters(10).Font.Subscript = True
```

This example checks the selected text for subscript formatting.




```vb
If Selection.Type = wdSelectionNormal Then 
 mySel = Selection.Font.Subscript 
 If mySel = wdUndefined Or mySel = True Then 
 MsgBox "Subscript text exists in the selection." 
 Else 
 MsgBox "No subscript text in the selection." 
 End If 
Else 
 MsgBox "You need to select some text." 
End If
```


## See also


#### Concepts


[Font Object](font-object-word.md)

