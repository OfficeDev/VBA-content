---
title: Section.Headers Property (Word)
keywords: vbawd10.chm156827769
f1_keywords:
- vbawd10.chm156827769
ms.prod: word
api_name:
- Word.Section.Headers
ms.assetid: 72b61449-2f93-a67a-2757-3c0441961307
ms.date: 06/08/2017
---


# Section.Headers Property (Word)

Returns a  **[HeadersFooters](headersfooters-object-word.md)** collection that represents the headers for the specified section. Read-only.


## Syntax

 _expression_ . **Headers**

 _expression_ A variable that represents a **[Section](section-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx). To return a  **HeadersFooters** collection that represents the footers for the specified section, use the **[Footers](section-footers-property-word.md)** property.


## Example

This example adds centered page numbers to every page in the active document except the first. (A separate header is created for the first page.)


```vb
With ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary) 
 .PageNumbers.Add _ 
 PageNumberAlignment:=wdAlignPageNumberCenter, _ 
 FirstPage:=False 
End With
```

This example adds text to the first-page header in the active document.




```vb
ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = True 
With ActiveDocument.Sections(1).Headers(wdHeaderFooterFirstPage) 
 .Range.InsertAfter("First Page Text") 
 .Range.Paragraphs.Alignment = wdAlignParagraphRight 
End With
```


## See also


#### Concepts


[Section Object](section-object-word.md)

