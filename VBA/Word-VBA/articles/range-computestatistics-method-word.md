---
title: Range.ComputeStatistics Method (Word)
keywords: vbawd10.chm157155506
f1_keywords:
- vbawd10.chm157155506
ms.prod: word
api_name:
- Word.Range.ComputeStatistics
ms.assetid: 5fbeeffd-f592-3078-cd5b-1e2a90ee5092
ms.date: 06/08/2017
---


# Range.ComputeStatistics Method (Word)

Returns a  **Long** that represents a statistic based on the contents of the specified range.


## Syntax

 _expression_ . **ComputeStatistics**( **_Statistic_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Statistic_|Required| **[WdStatistic](wdstatistic-enumeration-word.md)**|The type of statistic to compute.|

## Remarks

Some of the constants listed above may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example displays the number of words and characters in the first paragraph of Report.doc.


```vb
Set myRange = Documents("Report.doc").Paragraphs(1).Range 
wordCount = myRange.ComputeStatistics(Statistic:=wdStatisticWords) 
charCount = myRange.ComputeStatistics(Statistic:=wdStatisticCharacters) 
MsgBox "The first paragraph contains " &; wordCount _ 
 &; " words and a total of " &; charCount &; " characters."
```


## See also


#### Concepts


[Range Object](range-object-word.md)

