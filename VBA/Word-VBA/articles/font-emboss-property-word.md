---
title: Font.Emboss Property (Word)
keywords: vbawd10.chm156369044
f1_keywords:
- vbawd10.chm156369044
ms.prod: word
api_name:
- Word.Font.Emboss
ms.assetid: ae0cc2d0-b1ae-3208-7f61-cad731f04e29
ms.date: 06/08/2017
---


# Font.Emboss Property (Word)

 **True** if the specified font is formatted as embossed. Read/write **Long** .


## Syntax

 _expression_ . **Emboss**

 _expression_ A variable that represents a **[Font](font-object-word.md)** object.


## Remarks

Returns  **True** , **False** , or **wdUndefined** . Can be set to **True** , **False** , or **wdToggle** . Setting **Emboss** to **True** sets **[Engrave](font-engrave-property-word.md)** to **False** , and vice versa.


## Example

This example embosses the second sentence in a new document.


```vb
With Documents.Add.Content 
 .InsertAfter "This is the first sentence. " 
 .InsertAfter "This is the second sentence. " 
 .Sentences(2).Font.Emboss = True 
End With
```

This example embosses the selected text.




```vb
If Selection.Type = wdSelectionNormal Then 
 Selection.Font.Emboss = True 
Else 
 MsgBox "You need to select some text." 
End If
```


## See also


#### Concepts


[Font Object](font-object-word.md)

