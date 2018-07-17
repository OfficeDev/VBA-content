---
title: Range.TextRetrievalMode Property (Word)
keywords: vbawd10.chm157155390
f1_keywords:
- vbawd10.chm157155390
ms.prod: word
api_name:
- Word.Range.TextRetrievalMode
ms.assetid: e3992479-ba69-e8d3-17e3-73b533f27d26
ms.date: 06/08/2017
---


# Range.TextRetrievalMode Property (Word)

Returns a  **[TextRetrievalMode](textretrievalmode-object-word.md)** object that controls how text is retrieved from the specified **Range** . Read/write.


## Syntax

 _expression_ . **TextRetrievalMode**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Example

This example retrieves the selected text (excluding any hidden text) and inserts it at the beginning of the third paragraph in the active document.


```vb
If Selection.Type = wdSelectionNormal Then 
 Set Range1 = Selection.Range 
 Range1.TextRetrievalMode.IncludeHiddenText = False 
 Set Range2 = ActiveDocument.Paragraphs(2).Range 
 Range2.InsertAfter Range1.Text 
End If
```

This example retrieves and displays the first three paragraphs as they appear in outline view.




```vb
Set myRange = ActiveDocument.Range(Start:=ActiveDocument _ 
 .Paragraphs(1).Range.Start, _ 
 End:=ActiveDocument.Paragraphs(3).Range.End) 
myRange.TextRetrievalMode.ViewType = wdOutlineView 
MsgBox myRange.Text
```

This example excludes field codes and hidden text from the range that refers to the selected text. The example then displays the text in a message box.




```vb
If Selection.Type = wdSelectionNormal Then 
 Set aRange = Selection.Range 
 With aRange.TextRetrievalMode 
 .IncludeHiddenText = False 
 .IncludeFieldCodes = False 
 End With 
 MsgBox aRange.Text 
End If
```


## See also


#### Concepts


[Range Object](range-object-word.md)

