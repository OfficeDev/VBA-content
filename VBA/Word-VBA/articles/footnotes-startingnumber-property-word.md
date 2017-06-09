---
title: Footnotes.StartingNumber Property (Word)
keywords: vbawd10.chm155320422
f1_keywords:
- vbawd10.chm155320422
ms.prod: word
api_name:
- Word.Footnotes.StartingNumber
ms.assetid: c058fc5b-80d0-beba-5e89-2f8661264122
ms.date: 06/08/2017
---


# Footnotes.StartingNumber Property (Word)

Returns or sets the starting note number, line number, or page number. Read/write  **Long** .


## Syntax

 _expression_ . **StartingNumber**

 _expression_ Required. A variable that represents a **[Footnotes](footnotes-object-word.md)** collection.


## Remarks

You must be in print layout view to see line numbering.

When applied to page numbers, this property returns or sets the beginning page number for the specified  **HeaderFooter** object. This number may or may not be visible on the first page, depending on the setting of the **ShowFirstPageNumber** property. The **RestartNumberingAtSection** property, if set to **False** , will override the **StartingNumber** property so that page numbering can continue from the previous section.


## Example

This example creates a new document, sets the starting number for footnotes to 10, and then adds a footnote at the insertion point.


```vb
Set myDoc = Documents.Add 
With myDoc.Footnotes 
 .StartingNumber = 10 
 .Add Range:=Selection.Range, Text:="Text for a footnote" 
End With
```


## See also


#### Concepts


[Footnotes Collection Object](footnotes-object-word.md)

