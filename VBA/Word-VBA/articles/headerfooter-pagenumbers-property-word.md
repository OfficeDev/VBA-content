---
title: HeaderFooter.PageNumbers Property (Word)
keywords: vbawd10.chm159711237
f1_keywords:
- vbawd10.chm159711237
ms.prod: word
api_name:
- Word.HeaderFooter.PageNumbers
ms.assetid: 2e36c668-f696-e09e-dd04-ae77e7524232
ms.date: 06/08/2017
---


# HeaderFooter.PageNumbers Property (Word)

Returns a  **[PageNumbers](pagenumbers-object-word.md)** collection that represents all the page number fields included in the specified header or footer.


## Syntax

 _expression_ . **PageNumbers**

 _expression_ An expression that returns a **[HeaderFooter](headerfooter-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example creates a new document and adds page numbers to the footer.


```vb
Set myDoc = Documents.Add 
With myDoc.Sections(1).Footers(wdHeaderFooterPrimary) 
 .PageNumbers.Add PageNumberAlignment := wdAlignPageNumberCenter 
End With
```


## See also


#### Concepts


[HeaderFooter Object](headerfooter-object-word.md)

