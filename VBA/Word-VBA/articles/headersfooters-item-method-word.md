---
title: HeadersFooters.Item Method (Word)
keywords: vbawd10.chm159645696
f1_keywords:
- vbawd10.chm159645696
ms.prod: word
api_name:
- Word.HeadersFooters.Item
ms.assetid: b6449c22-d528-acce-4308-28fa81e596c5
ms.date: 06/08/2017
---


# HeadersFooters.Item Method (Word)

Returns a  **HeaderFooter** object that represents a header or footer in a range or section.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ Required. A variable that represents a **[HeadersFooters](headersfooters-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **WdHeaderFooterIndex**|A constant that specifies the header or footer in the range or section.|

### Return Value

HeaderFooter


## Example

This example creates and formats a first page header in the active document.


```vb
Sub HeadFootItem() 
 ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = True 
 With ActiveDocument.Sections(1).Headers _ 
 .Item(wdHeaderFooterFirstPage).Range 
 .InsertBefore "Sales Report" 
 With .Font 
 .Bold = True 
 .Size = "15" 
 .Color = wdColorBlue 
 End With 
 .Paragraphs.Alignment = wdAlignParagraphCenter 
 End With 
End Sub
```


## See also


#### Concepts


[HeadersFooters Collection Object](headersfooters-object-word.md)

