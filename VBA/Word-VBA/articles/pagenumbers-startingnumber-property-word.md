---
title: PageNumbers.StartingNumber Property (Word)
keywords: vbawd10.chm159776775
f1_keywords:
- vbawd10.chm159776775
ms.prod: word
api_name:
- Word.PageNumbers.StartingNumber
ms.assetid: 7b526500-2251-dedd-3972-52d4f036d3bc
ms.date: 06/08/2017
---


# PageNumbers.StartingNumber Property (Word)

Returns or sets the starting note number, line number, or page number. Read/write  **Long** .


## Syntax

 _expression_ . **StartingNumber**

 _expression_ Required. An expression that returns a **[PageNumbers](pagenumbers-object-word.md)** object.


## Remarks

You must be in print layout view to see line numbering.

When applied to page numbers, this property returns or sets the beginning page number for the specified  **[HeaderFooter](headerfooter-object-word.md)** object. This number may or may not be visible on the first page, depending on the setting of the **[ShowFirstPageNumber](pagenumbers-showfirstpagenumber-property-word.md)** property. The **[RestartNumberingAtSection](pagenumbers-restartnumberingatsection-property-word.md)** property, if set to **False** , will override the **StartingNumber** property so that page numbering can continue from the previous section.


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


[PageNumbers Collection Object](pagenumbers-object-word.md)

