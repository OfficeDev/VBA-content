---
title: PageNumbers.ShowFirstPageNumber Property (Word)
keywords: vbawd10.chm159776776
f1_keywords:
- vbawd10.chm159776776
ms.prod: word
api_name:
- Word.PageNumbers.ShowFirstPageNumber
ms.assetid: 5f7c88cc-ddb7-08d6-880d-f55a9591fdea
ms.date: 06/08/2017
---


# PageNumbers.ShowFirstPageNumber Property (Word)

 **True** if the page number appears on the first page in the section. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowFirstPageNumber**

 _expression_ An expression that returns a **[PageNumbers](pagenumbers-object-word.md)** collection object.


## Remarks

Setting this property to  **True** automatically adds page numbers to a section.


## Example

This example checks to see whether the page number appears on the first page in the active document.


```vb
Set myDoc = ActiveDocument 
first = myDoc.Sections(1).Headers(wdHeaderFooterPrimary). _ 
 PageNumbers.ShowFirstPageNumber 
Msgbox "This document shows numbers on the first page - " &; first
```

This example adds page numbers to the active document.




```vb
ActiveDocument.Sections(1) _ 
 .Headers(wdHeaderFooterPrimary).PageNumbers _ 
 .ShowFirstPageNumber = True
```


## See also


#### Concepts


[PageNumbers Collection Object](pagenumbers-object-word.md)

