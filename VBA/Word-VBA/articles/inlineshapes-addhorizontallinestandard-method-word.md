---
title: InlineShapes.AddHorizontalLineStandard Method (Word)
keywords: vbawd10.chm162070633
f1_keywords:
- vbawd10.chm162070633
ms.prod: word
api_name:
- Word.InlineShapes.AddHorizontalLineStandard
ms.assetid: de9d4613-4e64-9df8-aa9a-890335eb648d
ms.date: 06/08/2017
---


# InlineShapes.AddHorizontalLineStandard Method (Word)

Adds a horizontal line to the current document.


## Syntax

 _expression_ . **AddHorizontalLineStandard**( **_Range_** )

 _expression_ Required. A variable that represents an **[InlineShapes](inlineshapes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Optional| **Variant**|The range above which Microsoft Word places the horizontal line. If this argument is omitted, Word places the horizontal line above the current selection.|

## Remarks

To add a horizontal line based on an existing image file, use the  **[AddHorizontalLine](inlineshapes-addhorizontalline-method-word.md)** method.


## Example

This example adds a horizontal line above the fifth paragraph in the active document.


```vb
ActiveDocument.Paragraphs(5).Range _ 
 .InlineShapes.AddHorizontalLineStandard
```


## See also


#### Concepts


[InlineShapes Collection Object](inlineshapes-object-word.md)

