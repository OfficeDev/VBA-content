---
title: PageNumbers.Add Method (Word)
keywords: vbawd10.chm159776869
f1_keywords:
- vbawd10.chm159776869
ms.prod: word
api_name:
- Word.PageNumbers.Add
ms.assetid: d8a81795-035b-9702-bcd4-02c302607670
ms.date: 06/08/2017
---


# PageNumbers.Add Method (Word)

Returns a  **PageNumber** object that represents page numbers added to a header or footer in a section.


## Syntax

 _expression_ . **Add**( **_PageNumberAlignment_** , **_FirstPage_** )

 _expression_ Required. A variable that represents a **[PageNumbers](pagenumbers-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PageNumberAlignment_|Optional| **Variant**|Can be any  **WdPageNumberAlignment** constant.|
| _FirstPage_|Optional| **Variant**| **False** to make the first-page header and the first-page footer different from the headers and footers on all subsequent pages in the document. If FirstPage is set to **False** , a page number isn't added to the first page. If this argument is omitted, the setting is controlled by the **DifferentFirstPageHeaderFooter** property.|

## Remarks

If the  **LinkToPrevious** property for the **HeaderFooter** object is set to **True** , the page numbers will continue sequentially from one section to next throughout the document.


## Example

This example adds a page number to the primary footer in the first section of the active document.


```vb
With ActiveDocument.Sections(1) 
 .Footers(wdHeaderFooterPrimary).PageNumbers.Add _ 
 PageNumberAlignment:=wdAlignPageNumberLeft, _ 
 FirstPage:=True 
End With
```

This example creates and formats page numbers in the header for the active document.




```vb
Set myPgNum = ActiveDocument.Sections(1) _ 
 .Headers(wdHeaderFooterPrimary) _ 
 .PageNumbers.Add(PageNumberAlignment:= _ 
 wdAlignPageNumberCenter, FirstPage:= True) 
myPgNum.Select 
With Selection.Range 
 .Italic = True 
 .Bold = True 
End With
```


## See also


#### Concepts


[PageNumbers Collection Object](pagenumbers-object-word.md)

