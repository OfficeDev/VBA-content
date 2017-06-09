---
title: Font.Engrave Property (Word)
keywords: vbawd10.chm156369046
f1_keywords:
- vbawd10.chm156369046
ms.prod: word
api_name:
- Word.Font.Engrave
ms.assetid: 9d062637-05c8-d1c9-2231-23439bed30b9
ms.date: 06/08/2017
---


# Font.Engrave Property (Word)

 **True** if the font is formatted as engraved. Read/write **Long** .


## Syntax

 _expression_ . **Engrave**

 _expression_ A variable that represents a **[Font](font-object-word.md)** object.


## Remarks

Returns  **True** , **False** or **wdUndefined** (a mixture of **True** and **False** ). Can be set to **True** , **False** , or **wdToggle** . Setting **Engrave** to **True** sets **[Emboss](font-emboss-property-word.md)** to **False** , and vice versa.


## Example

This example formats the first letter in the active document as engraved.


```vb
Dim rngTemp As Range 
 
Set rngTemp = ActiveDocument.Characters(1) 
With rngTemp.Font 
 .Size = 20 
 .Engrave = True 
End With
```

This example formats the selection as engraved.




```vb
If Selection.Type = wdSelectionNormal Then 
 Selection.Font.Engrave = True 
Else 
 MsgBox "You need to select some text." 
End If
```


## See also


#### Concepts


[Font Object](font-object-word.md)

