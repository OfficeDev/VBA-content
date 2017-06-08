---
title: Font.Italic Property (Word)
keywords: vbawd10.chm156369027
f1_keywords:
- vbawd10.chm156369027
ms.prod: word
api_name:
- Word.Font.Italic
ms.assetid: adba2e3c-d904-d835-5a1c-c8762d319106
ms.date: 06/08/2017
---


# Font.Italic Property (Word)

 **True** if the font or range is formatted as italic. Read/write **Long** .


## Syntax

 _expression_ . **Italic**

 _expression_ A variable that represents a **[Font](font-object-word.md)** object.


## Remarks

This property returns  **True** , **False** or **wdUndefined** (a mixture of **True** and **False** ). Can be set to **True** , **False** , or **wdToggle** .


## Example

This example checks the selection for italic formatting and removes any that it finds.


```vb
If Selection.Type = wdSelectionNormal Then 
 mySel = Selection.Font.Italic 
 If mySel = wdUndefined or mySel = True Then 
 MsgBox "there is italic text in selection. " _ 
 &; "Click OK to remove." 
 Selection.Font.Italic = False 
 Else 
 MsgBox "No italic text in the selection." 
 End If 
Else 
 MsgBox "You need to select some text." 
End If
```


## See also


#### Concepts


[Font Object](font-object-word.md)

