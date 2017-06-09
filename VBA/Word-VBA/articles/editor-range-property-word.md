---
title: Editor.Range Property (Word)
keywords: vbawd10.chm225575014
f1_keywords:
- vbawd10.chm225575014
ms.prod: word
api_name:
- Word.Editor.Range
ms.assetid: a09abe23-cc64-2fda-682d-7d2825a9e5fb
ms.date: 06/08/2017
---


# Editor.Range Property (Word)

Returns a  **Range** object that represents the portion of a document that is contained in the specified object.


## Syntax

 _expression_ . **Range**

 _expression_ Required. A variable that represents an **[Editor](editor-object-word.md)** object.


## Remarks

For information about returning a range from a document or returning a shape range from a collection of shapes, see the  **Range** method.


## Example

This example applies the Heading 1 style to the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).Range.Style = wdStyleHeading1
```

This example copies the first row in table one.




```vb
If ActiveDocument.Tables.Count >= 1 Then _ 
 ActiveDocument.Tables(1).Rows(1).Range.Copy
```

This example changes the text of the first comment in the document.




```vb
With ActiveDocument.Comments(1).Range 
 .Delete 
 .InsertBefore "new comment text" 
End With
```

This example inserts text at the end of section one.




```vb
Set myRange = ActiveDocument.Sections(1).Range 
With myRange 
 .MoveEnd Unit:=wdCharacter, Count:=-1 
 .Collapse Direction:=wdCollapseEnd 
 .InsertParagraphAfter 
 .InsertAfter "End of section" 
End With
```


## See also


#### Concepts


[Editor Object](editor-object-word.md)

