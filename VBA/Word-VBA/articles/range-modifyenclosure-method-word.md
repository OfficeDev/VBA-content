---
title: Range.ModifyEnclosure Method (Word)
keywords: vbawd10.chm157155551
f1_keywords:
- vbawd10.chm157155551
ms.prod: word
api_name:
- Word.Range.ModifyEnclosure
ms.assetid: 173c5b41-5245-4fc5-b9d9-9fd7cea0aab8
ms.date: 06/08/2017
---


# Range.ModifyEnclosure Method (Word)

Adds, modifies, or removes an enclosure around the specified character or characters.


## Syntax

 _expression_ . **ModifyEnclosure**( **_Style_** , **_Symbol_** , **_EnclosedText_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **Variant**|The style of the enclosure. Can be any  **WdEncloseStyle** constant.|
| _Symbol_|Optional| **Variant**|The symbol in which to enclose the specified range. Can be any  **WdEnclosureType** constant.|
| _EnclosedText_|Optional| **Variant**|The characters that you want to enclose. If you include this argument, Microsoft Word replaces the specified range with the enclosed characters. If you don't specify text to enclose, Microsoft Word encloses all text in the specified range.|

## Example

This example replaces the current selection with the number 25 enclosed in a circle.


```
Selection.Range.ModifyEnclosure wdEncloseStyleLarge, _ 
 wdEnclosureCircle, "25"
```


## See also


#### Concepts


[Range Object](range-object-word.md)

