---
title: Range.GoToNext Method (Word)
keywords: vbawd10.chm157155502
f1_keywords:
- vbawd10.chm157155502
ms.prod: word
api_name:
- Word.Range.GoToNext
ms.assetid: 011de2d6-c0fc-608f-8d7e-faac5947978d
ms.date: 06/08/2017
---


# Range.GoToNext Method (Word)

Returns a  **Range** object that refers to the start position of the next item or location specified by the What argument. .


## Syntax

 _expression_ . **GoToNext**( **_What_** )

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

