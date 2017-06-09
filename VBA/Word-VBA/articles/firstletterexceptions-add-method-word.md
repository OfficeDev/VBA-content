---
title: FirstLetterExceptions.Add Method (Word)
keywords: vbawd10.chm155582565
f1_keywords:
- vbawd10.chm155582565
ms.prod: word
api_name:
- Word.FirstLetterExceptions.Add
ms.assetid: 66ed8423-2c64-e924-2b34-45daea68efac
ms.date: 06/08/2017
---


# FirstLetterExceptions.Add Method (Word)

Returns a  **FirstLetterException** object that represents a new exception added to the list of AutoCorrect exceptions.


## Syntax

 _expression_ . **Add**( **_Name_** )

 _expression_ Required. A variable that represents a **[FirstLetterExceptions](firstletterexceptions-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The word with two initial capital letters that you want Microsoft Word to overlook.|

### Return Value

FirstLetterException


## Remarks

If the  **FirstLetterAutoAdd** property is **True** , abbreviations are automatically added to the list of first-letter exceptions.


## Example

This example adds the abbreviation addr. to the list of first-letter exceptions.


```
AutoCorrect.FirstLetterExceptions.Add Name:="addr."
```


## See also


#### Concepts


[FirstLetterExceptions Collection Object](firstletterexceptions-object-word.md)

