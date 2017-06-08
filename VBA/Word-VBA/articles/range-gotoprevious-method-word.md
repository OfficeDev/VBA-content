---
title: Range.GoToPrevious Method (Word)
keywords: vbawd10.chm157155503
f1_keywords:
- vbawd10.chm157155503
ms.prod: word
api_name:
- Word.Range.GoToPrevious
ms.assetid: b1a6d089-c36a-1e10-fd8e-090d5b736a88
ms.date: 06/08/2017
---


# Range.GoToPrevious Method (Word)

Returns a  **Range** object that refers to the start position of the previous item or location specified by the What argument.


## Syntax

 _expression_ . **GoToPrevious**( **_What_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _What_|Required| **[WdGoToItem](wdgotoitem-enumeration-word.md)**|The item to where the specified range or selection is to be moved.|

## Remarks




 **Note**  When you use this method with the  **wdGoToGrammaticalError** , **wdGoToProofreadingError** , or **wdGoToSpellingError** constant, the **Range** object that's returned includes any grammar error text or spelling error text.


## See also


#### Concepts


[Range Object](range-object-word.md)

