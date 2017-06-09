---
title: Index.IndexLanguage Property (Word)
keywords: vbawd10.chm159186954
f1_keywords:
- vbawd10.chm159186954
ms.prod: word
api_name:
- Word.Index.IndexLanguage
ms.assetid: 1fcc2332-eba2-ee2d-67ea-f256254d3c2c
ms.date: 06/08/2017
---


# Index.IndexLanguage Property (Word)

Returns or sets a  **[WdLanguageID](wdlanguageid-enumeration-word.md)** constant that represents the sorting language to use for the specified index. Read/write .


## Syntax

 _expression_**IndexLanguage**

 _expression_ Required. An expression that returns an **[Index](index-object-word.md)** object.


## Remarks

Some of the  **WdLanguageID** constants may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example sets the sorting language of the first index in the active document to New Zealand English.


```vb
ActiveDocument.Indexes(1).IndexLanguage = _ 
 wdEnglishNewZealand
```


## See also


#### Concepts


[Index Object](index-object-word.md)

