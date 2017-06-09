---
title: Font.Size Property (Word)
keywords: vbawd10.chm156369037
f1_keywords:
- vbawd10.chm156369037
ms.prod: word
api_name:
- Word.Font.Size
ms.assetid: 95041c0a-43d0-368c-1c75-e2d6c1821119
ms.date: 06/08/2017
---


# Font.Size Property (Word)

Returns or sets the font size, in points. Read/write  **Single** .


## Syntax

 _expression_ . **Size**

 _expression_ Required. A variable that represents a **[Font](font-object-word.md)** object.


## Example

This example inserts text and then sets the font size of the seventh word of the inserted text to 20 points.


```vb
Selection.Collapse Direction:=wdCollapseEnd 
With Selection.Range 
 .Font.Reset 
 .InsertBefore "This is a demonstration of font size." 
 .Words(7).Font.Size = 20 
End With
```

This example determines the font size of the selected text.




```vb
mySel = Selection.Font.Size 
If mySel = wdUndefined Then 
 MsgBox "there is a mix of font sizes in the selection." 
Else 
 MsgBox mySel &; " points" 
End If
```


## See also


#### Concepts


[Font Object](font-object-word.md)

