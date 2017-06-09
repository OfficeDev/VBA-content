---
title: Font.AllCaps Property (Word)
keywords: vbawd10.chm156369030
f1_keywords:
- vbawd10.chm156369030
ms.prod: word
api_name:
- Word.Font.AllCaps
ms.assetid: ef881fd6-bb35-7cc6-b048-c9ed2111f821
ms.date: 06/08/2017
---


# Font.AllCaps Property (Word)

 **True** if the font is formatted as all capital letters. Read/write **Long** .


## Syntax

 _expression_ . **AllCaps**

 _expression_ A variable that represents a **[Font](font-object-word.md)** object.


## Remarks

Returns  **True** , **False** , or wdUndefined (a mixture of **True** and **False** ). Can be set to **True** , **False** , or **wdToggle** (reverses the current setting).

Setting  **AllCaps** to **True** sets **[SmallCaps](font-smallcaps-property-word.md)** to **False** , and vice versa.


## Example

This example checks the third paragraph in the active document for text formatted as all capital letters.


```vb
If ActiveDocument.Paragraphs(3).Range.Font.AllCaps = True Then 
 MsgBox "Text is all caps." 
Else 
 MsgBox "Text is not all caps." 
End if
```

This example formats the selected text as all capital letters.




```vb
If Selection.Type = wdSelectionNormal Then 
 Selection.Font.AllCaps = True 
Else 
 MsgBox "You need to select some text." 
End If
```


## See also


#### Concepts


[Font Object](font-object-word.md)

