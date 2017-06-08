---
title: CanvasShapes.Count Property (Word)
keywords: vbawd10.chm7536642
f1_keywords:
- vbawd10.chm7536642
ms.prod: word
api_name:
- Word.CanvasShapes.Count
ms.assetid: d6f54f95-716b-1b6a-33b8-0dbbc1006a2b
ms.date: 06/08/2017
---


# CanvasShapes.Count Property (Word)

Returns a  **Long** that represents the number of canvas shapes in the specified collection. Read-only.


## Syntax

 _expression_ . **Count**

 _expression_ Required. A variable that represents a **[CanvasShapes](canvasshapes-object-word.md)** collection.


## Example

This example displays the number of paragraphs in the active document.


```vb
MsgBox "The active document contains " &; _ 
 ActiveDocument.Paragraphs.Count &; " paragraphs."
```

This example displays the number of words in the selection.




```vb
If Selection.Words.Count >= 1 And _ 
 Selection.Type <> wdSelectionIP Then 
 MsgBox "The selection contains " &; Selection.Words.Count _ 
 &; " words." 
End If
```

This example uses the aFields() array to store the field codes in the active document.




```vb
fcount = ActiveDocument.Fields.Count 
If fcount >= 1 Then 
 ReDim aFields(fcount) 
 For Each aField In ActiveDocument.Fields 
 aFields(aField.Index) = aField.Code.Text 
 Next aField 
End If
```


## See also


#### Concepts


[CanvasShapes Collection](canvasshapes-object-word.md)

