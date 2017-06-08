---
title: PageNumber Object (Word)
ms.prod: word
api_name:
- Word.PageNumber
ms.assetid: 5b58f562-de19-ac9d-0b2c-7696603c1606
ms.date: 06/08/2017
---


# PageNumber Object (Word)

Represents a page number in a header or footer. The  **PageNumber** object is a member of the **[PageNumbers](pagenumbers-object-word.md)** collection. The **PageNumbers** collection includes all the page numbers in a single header or footer.


## Remarks

Use  **PageNumbers** (Index), where Index is the index number, to return a single **PageNumber** object. In most cases, a header or footer will contain only one page number, which is index number 1. The following example centers the first page number in the primary header in section one in the active document.


```vb
ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary) _ 
 .PageNumbers(1).Alignment = wdAlignPageNumberCenter
```

Use the  **Add** method to add a page number (a PAGE field) to a header or footer. The following example adds a page number to the primary footer in the first section and in any subsequent sections. The page number doesn't appear on the first page.




```vb
With ActiveDocument.Sections(1) 
 .Footers(wdHeaderFooterPrimary).PageNumbers.Add _ 
 PageNumberAlignment:=wdAlignPageNumberLeft, _ 
 FirstPage:=False 
End With
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

