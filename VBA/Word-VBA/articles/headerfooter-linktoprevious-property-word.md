---
title: HeaderFooter.LinkToPrevious Property (Word)
keywords: vbawd10.chm159711238
f1_keywords:
- vbawd10.chm159711238
ms.prod: word
api_name:
- Word.HeaderFooter.LinkToPrevious
ms.assetid: edb4dbeb-bb06-e41a-aa26-f29317bb6e01
ms.date: 06/08/2017
---


# HeaderFooter.LinkToPrevious Property (Word)

 **True** if the specified header or footer is linked to the corresponding header or footer in the previous section. Read/write **Boolean** .


## Syntax

 _expression_ . **LinkToPrevious**

 _expression_ An expression that returns a **[HeaderFooter](headerfooter-object-word.md)** object.


## Remarks

When a header or footer is linked, its contents are the same as in the previous header or footer. Because the  **LinkToPrevious** property is set to **True** by default, you can add headers, footers, and page numbers to your entire document by working with the headers, footers, and page numbers in the first section. For instance, the following example adds page numbers to the header on all pages in all sections of the active document.


```vb
ActiveDocument.Sections(1) _ 
 .Headers(wdHeaderFooterPrimary).PageNumbers.Add
```

The  **LinkToPrevious** property applies to each header or footer individually. For example, the **LinkToPrevious** property could be set to **True** for the even-numbered-page header but **False** for the even-numbered-page footer.


## Example

The first part of this example creates a new document with two sections. The second part creates unique headers for even-numbered and odd-numbered pages in sections one and two in the new document.


```vb
Documents.Add 
With Selection 
 For j = 1 to 4 
 .TypeParagraph 
 .InsertBreak 
 .TypeParagraph 
 Next j 
End With 
With ActiveDocument 
 .Paragraphs(5).Range.InsertBreak Type:=wdSectionBreakNextPage 
 .PageSetup.OddAndEvenPagesHeaderFooter = True 
End With 
With ActiveDocument.Sections(2) 
 With .Headers(wdHeaderFooterPrimary) 
 .LinkToPrevious = False 
 .Range.InsertBefore "Section 2 Odd Header" 
 End With 
 With .Headers(wdHeaderFooterEvenPages) 
 .LinkToPrevious = False 
 .Range.InsertBefore "Section 2 Even Header" 
 End With 
End With 
With ActiveDocument.Sections(1) 
 .Headers(wdHeaderFooterPrimary) _ 
 .Range.InsertBefore "Section 1 Odd Header" 
 .Headers(wdHeaderFooterEvenPages) _ 
 .Range.InsertBefore "Section 1 Even Header" 
End With
```


## See also


#### Concepts


[HeaderFooter Object](headerfooter-object-word.md)

