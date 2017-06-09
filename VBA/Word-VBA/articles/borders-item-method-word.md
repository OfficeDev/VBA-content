---
title: Borders.Item Method (Word)
keywords: vbawd10.chm154927104
f1_keywords:
- vbawd10.chm154927104
ms.prod: word
api_name:
- Word.Borders.Item
ms.assetid: ac2b9108-5ae1-e875-f6a0-47a8c2175fe1
ms.date: 06/08/2017
---


# Borders.Item Method (Word)

Returns a border in a range or selection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ Required. A variable that represents a **[Borders](borders-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **WdBorderType**|The border to be returned.|

### Return Value

Border


## Example

This example inserts a double border above the first paragraph in the active document.


```vb
Sub BorderItem() 
 ActiveDocument.Paragraphs(1).Borders.Item(wdBorderTop) _ 
 .LineStyle = wdLineStyleDouble 
End Sub
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

