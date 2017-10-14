---
title: LineNumbering.StartingNumber Property (Word)
keywords: vbawd10.chm158466149
f1_keywords:
- vbawd10.chm158466149
ms.prod: word
api_name:
- Word.LineNumbering.StartingNumber
ms.assetid: 115d4c0a-d895-a404-84bb-7ffe17706a98
ms.date: 06/08/2017
---


# LineNumbering.StartingNumber Property (Word)

Returns or sets the starting line number. Read/write  **Long** .


## Syntax

 _expression_ . **StartingNumber**

 _expression_ Required. A variable that represents a **[LineNumbering](linenumbering-object-word.md)** object.


## Remarks

You must be in print layout view to see line numbering.


## Example

This example creates a new document, sets the starting number for footnotes to 10, and then adds a footnote at the insertion point.


```vb
Set myDoc = Documents.Add 
With myDoc.Footnotes 
 .StartingNumber = 10 
 .Add Range:=Selection.Range, Text:="Text for a footnote" 
End With
```

This example enables line numbering for the active document. The starting number is set to 5, every fifth line number is shown, and the numbering starts over at the beginning of each section in the document.




```vb
With ActiveDocument.PageSetup.LineNumbering 
 .Active = True 
 .StartingNumber = 5 
 .CountBy = 5 
 .RestartMode = wdRestartSection 
End With
```

This example sets properties for page numbers, and then it adds page numbers to the header of the active document.




```vb
With ActiveDocument.Sections(1) _ 
 .Headers(wdHeaderFooterPrimary).PageNumbers 
 .NumberStyle = wdPageNumberStyleArabic 
 .IncludeChapterNumber = False 
 .RestartNumberingAtSection = True 
 .StartingNumber = 5 
 .Add PageNumberAlignment:=wdAlignPageNumberCenter, _ 
 FirstPage:=True 
End With
```


## See also


#### Concepts


[LineNumbering Object](linenumbering-object-word.md)

