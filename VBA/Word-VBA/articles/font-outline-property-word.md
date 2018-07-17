---
title: Font.Outline Property (Word)
keywords: vbawd10.chm156369043
f1_keywords:
- vbawd10.chm156369043
ms.prod: word
api_name:
- Word.Font.Outline
ms.assetid: f2ec3056-5b5d-be3c-af8d-1eed86b4d01e
ms.date: 06/08/2017
---


# Font.Outline Property (Word)

 **True** if the font is formatted as outline. Read/write **Long** .


## Syntax

 _expression_ . **Outline**

 _expression_ An expression that returns a **[Font](font-object-word.md)** object.


## Remarks

Returns  **True** , **False** , or **wdUndefined** (a mixture of **True** and **False** ). Can be set to **True** , **False** , or **wdToggle** .


## Example

This example applies outline font formatting to the first three words in the active document.


```vb
Set myRange = ActiveDocument.Range(Start:= _ 
 ActiveDocument.Words(1).Start, _ 
 End:=ActiveDocument.Words(3).End) 
myRange.Font.Outline = True
```

This example toggles outline formatting for the selected text.




```
Selection.Font.Outline = wdToggle
```

This example removes outline font formatting from the selection if outline formatting is partially applied to the selection.




```vb
Set myFont = Selection.Font 
If myFont.Outline = wdUndefined Then 
 myFont.Outline = False 
End If
```


## See also


#### Concepts


[Font Object](font-object-word.md)

