---
title: Font.Spacing Property (Word)
keywords: vbawd10.chm156369040
f1_keywords:
- vbawd10.chm156369040
ms.prod: word
api_name:
- Word.Font.Spacing
ms.assetid: 50e380cd-1126-c2bd-18ff-40249efa0b9f
ms.date: 06/08/2017
---


# Font.Spacing Property (Word)

Returns or sets the spacing (in points) between characters. Read/write  **Single** .


## Syntax

 _expression_ . **Spacing**

 _expression_ Required. A variable that represents a **[Font](font-object-word.md)** object.


## Example

This example demonstrates two different character spacings at the beginning of the active document.


```vb
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
With myRange 
 .InsertAfter "Demonstration of no character spacing." 
 .InsertParagraphAfter 
 .InsertAfter "Demonstration of character spacing (1.5pt)." 
 .InsertParagraphAfter 
End With 
ActiveDocument.Paragraphs(2).Range.Font.Spacing = 1.5
```

This example sets the character spacing of the selected text to 2 points.




```vb
If Selection.Type = wdSelectionNormal Then 
 Selection.Font.Spacing = 2 
Else 
 MsgBox "You need to select some text." 
End If
```


## See also


#### Concepts


[Font Object](font-object-word.md)

