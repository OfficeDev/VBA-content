---
title: Selection.Text Property (Word)
keywords: vbawd10.chm158662656
f1_keywords:
- vbawd10.chm158662656
ms.prod: word
api_name:
- Word.Selection.Text
ms.assetid: 2acf885b-8d4a-7ebc-79aa-902921bc33bb
ms.date: 06/08/2017
---


# Selection.Text Property (Word)

Returns or sets the text in the specified selection. Read/write  **String** .


## Syntax

 _expression_ . **Text**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

The  **Text** property returns the plain, unformatted text of the selection. When you set this property, the text of the range or selection is replaced.


## Example

This example displays the text in the selection. If nothing is selected, the character following the insertion point is displayed.


```vb
MsgBox Selection.Text
```

This example inserts 10 lines of text into a new document.




```vb
Documents.Add 
For i = 1 To 10 
 Selection.Text = "Line" &; Str(i) &; Chr(13) 
 Selection.MoveDown Unit:=wdParagraph, Count:=1 
Next i
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

