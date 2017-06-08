---
title: Document.CountNumberedItems Method (Word)
keywords: vbawd10.chm158007438
f1_keywords:
- vbawd10.chm158007438
ms.prod: word
api_name:
- Word.Document.CountNumberedItems
ms.assetid: b35face4-9d35-2071-90e1-628e7eca04fc
ms.date: 06/08/2017
---


# Document.CountNumberedItems Method (Word)

Returns the number of bulleted or numbered items and LISTNUM fields in the specified  **Document** object.


## Syntax

 _expression_ . **CountNumberedItems**( **_NumberType_** , **_Level_** )

 _expression_ An expression that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NumberType_|Optional| **Variant**|The type of numbers to be counted. Can be one of the  **WdNumberType** constants. The default value is **wdNumberAllNumbers** .|
| _Level_|Optional| **Variant**|A number that corresponds to the numbering level you want to count. If this argument is omitted, all levels are counted.|

## Remarks

Bulleted items are counted when either  **wdNumberParagraph** or **wdNumberAllNumbers** (the default) is specified for NumberType.

There are two types of numbers: preset numbers ( **wdNumberParagraph** ), which you can add to paragraphs by selecting a template in the **Bullets and Numbering** dialog box; and LISTNUM fields ( **wdNumberListNum** ), which allow you to add more than one number per paragraph.


## See also


#### Concepts


[Document Object](document-object-word.md)

