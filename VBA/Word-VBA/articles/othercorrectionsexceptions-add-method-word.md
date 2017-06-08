---
title: OtherCorrectionsExceptions.Add Method (Word)
keywords: vbawd10.chm165609573
f1_keywords:
- vbawd10.chm165609573
ms.prod: word
api_name:
- Word.OtherCorrectionsExceptions.Add
ms.assetid: 0bdb30c5-72f0-3dae-e0c5-b2ea48157626
ms.date: 06/08/2017
---


# OtherCorrectionsExceptions.Add Method (Word)

Returns an  **OtherCorrectionsException** object that represents a new exception added to the list of AutoCorrect exceptions.


## Syntax

 _expression_ . **Add**( **_Name_** )

 _expression_ Required. A variable that represents an **[OtherCorrectionsExceptions](othercorrectionsexceptions-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The word that you want Word to overlook.|

### Return Value

OtherCorrectionsException


## Remarks

If the  **OtherCorrectionsAutoAdd** property is **True** , words are automatically added to the list of other corrections exceptions.


## Example

This example adds myCompany to the list of other corrections exceptions.


```
AutoCorrect.OtherCorrectionsExceptions.Add Name:="myCompany"
```


## See also


#### Concepts


[OtherCorrectionsExceptions Collection Object](othercorrectionsexceptions-object-word.md)

