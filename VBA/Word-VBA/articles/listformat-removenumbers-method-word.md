---
title: ListFormat.RemoveNumbers Method (Word)
keywords: vbawd10.chm163578041
f1_keywords:
- vbawd10.chm163578041
ms.prod: word
api_name:
- Word.ListFormat.RemoveNumbers
ms.assetid: 80c0e408-683d-4639-733d-843d5fd323e2
ms.date: 06/08/2017
---


# ListFormat.RemoveNumbers Method (Word)

Removes numbers or bullets from the specified list.


## Syntax

 _expression_ . **RemoveNumbers**( **_NumberType_** )

 _expression_ A variable that represents a **[ListFormat](listformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NumberType_|Optional| **[WdNumberType](wdnumbertype-enumeration-word.md)**| The type of number to be removed.|

## Example

This example removes the bullets or numbers from any numbered paragraphs in the selection.


```
Selection.Range.ListFormat.RemoveNumbers
```

This example removes the LISTNUM fields from the selection.




```
Selection.Range.ListFormat.RemoveNumbers wdNumberListNum
```


## See also


#### Concepts


[ListFormat Object](listformat-object-word.md)

