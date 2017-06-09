---
title: Font.Grow Method (Word)
keywords: vbawd10.chm156368996
f1_keywords:
- vbawd10.chm156368996
ms.prod: word
api_name:
- Word.Font.Grow
ms.assetid: 0bce9195-07df-d604-9208-1b1222a81b3e
ms.date: 06/08/2017
---


# Font.Grow Method (Word)

Increases the font size to the next available size.


## Syntax

 _expression_ . **Grow**

 _expression_ Required. A variable that represents a **[Font](font-object-word.md)** object.


## Remarks

If the selection or range contains more than one font size, each size is increased to the next available setting.


## Example

This example increases the font size of the fourth word in a new document.


```vb
Dim rngTemp As Range 
 
Set rngTemp = Documents.Add.Content 
rngTemp.InsertAfter "This is a test of the Grow method." 
MsgBox "Click OK to increase the font size of the fourth word." 
rngTemp.Words(4).Font.Grow
```

This example increases the font size of the selected text.




```vb
If Selection.Type = wdSelectionNormal Then 
 Selection.Font.Grow 
Else 
 MsgBox "You need to select some text." 
End If
```


## See also


#### Concepts


[Font Object](font-object-word.md)

