---
title: HangulAndAlphabetExceptions.Add Method (Word)
keywords: vbawd10.chm164692069
f1_keywords:
- vbawd10.chm164692069
ms.prod: word
api_name:
- Word.HangulAndAlphabetExceptions.Add
ms.assetid: 6cbfb762-4e14-a31a-1619-e8ad725b58c8
ms.date: 06/08/2017
---


# HangulAndAlphabetExceptions.Add Method (Word)

Returns a  **HangulAndAlphabetException** object that represents a new exception to the list of AutoCorrect exceptions.


## Syntax

 _expression_ . **Add**( **_Name_** )

 _expression_ Required. A variable that represents a **[HangulAndAlphabetExceptions](hangulandalphabetexceptions-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The word that you don't want Microsoft Word to correct automatically.|

### Return Value

HangulAndAlphabetException


## Remarks

If the  **HangulAndAlphabetAutoAdd** property is set to **True** , words are automatically added to the list of hangul and alphabet AutoCorrect exceptions.

For more information on using Word with East Asian languages, see Word features for East Asian languages .


## Example

This example adds test to the list of hangul and alphabet AutoCorrect exceptions on the  **Korean** tab in the **AutoCorrect Exceptions** dialog box.


```
AutoCorrect.HangulAndAlphabetExceptions.Add Name:="test"
```


## See also


#### Concepts


[HangulAndAlphabetExceptions Collection Object](hangulandalphabetexceptions-object-word.md)

