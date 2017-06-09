---
title: PageNumbers Object (Word)
ms.prod: word
ms.assetid: 9090f96e-d898-ace6-35fa-f6e59c527ea2
ms.date: 06/08/2017
---


# PageNumbers Object (Word)

A collection of  **PageNumber** objects that represent the page numbers in a single header or footer.


## Remarks

Use the  **PageNumbers** property to return the **PageNumbers** collection. The following example starts page numbering at 3 for the first section in the active document.


```
ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary) _ 
 .PageNumbers.StartingNumber = 3
```

Use the  **Add** method to add page numbers to a header or footer. The following example adds a page number to the primary footer in the first section.




```
With ActiveDocument.Sections(1) 
 .Footers(wdHeaderFooterPrimary).PageNumbers.Add _ 
 PageNumberAlignment:=wdAlignPageNumberLeft, _ 
 FirstPage:=False 
End With
```

To add or change page numbers in a document with multiple sections, modify the page numbers in each section or set the  **LinkToPrevious** property to **True**.

Use  **PageNumbers** (index), where index is the index number, to return a single **PageNumber** object. In most cases, a header or footer contains only one page number, which is index number 1. The following example centers the first page number in the primary header in the first section.




```
ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary) _ 
 .PageNumbers(1).Alignment = wdAlignPageNumberCenter
```


## Methods



|**Name**|
|:-----|
|[Add](pagenumbers-add-method-word.md)|
|[Item](pagenumbers-item-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](pagenumbers-application-property-word.md)|
|[ChapterPageSeparator](pagenumbers-chapterpageseparator-property-word.md)|
|[Count](pagenumbers-count-property-word.md)|
|[Creator](pagenumbers-creator-property-word.md)|
|[DoubleQuote](pagenumbers-doublequote-property-word.md)|
|[HeadingLevelForChapter](pagenumbers-headinglevelforchapter-property-word.md)|
|[IncludeChapterNumber](pagenumbers-includechapternumber-property-word.md)|
|[NumberStyle](pagenumbers-numberstyle-property-word.md)|
|[Parent](pagenumbers-parent-property-word.md)|
|[RestartNumberingAtSection](pagenumbers-restartnumberingatsection-property-word.md)|
|[ShowFirstPageNumber](pagenumbers-showfirstpagenumber-property-word.md)|
|[StartingNumber](pagenumbers-startingnumber-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
