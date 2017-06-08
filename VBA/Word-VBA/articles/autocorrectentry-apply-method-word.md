---
title: AutoCorrectEntry.Apply Method (Word)
keywords: vbawd10.chm155648102
f1_keywords:
- vbawd10.chm155648102
ms.prod: word
api_name:
- Word.AutoCorrectEntry.Apply
ms.assetid: 9427d4a3-e955-7fc6-eec2-d4580e95b13f
ms.date: 06/08/2017
---


# AutoCorrectEntry.Apply Method (Word)

Replaces a range with the value of the specified AutoCorrect entry.


## Syntax

 _expression_ . **Apply**( **_Range_** )

 _expression_ Required. A variable that represents an **[AutoCorrectEntry](autocorrectentry-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **[Range](range-object-word.md)**|The range to which to apply the AutoCorrect entry.|

## Example

This example adds an AutoCorrect replacement entry, then applies the "sr" AutoCorrect entry to the selected text.


```
AutoCorrect.Entries.Add Name:= "sr", Value:= "Stella Richards" 
AutoCorrect.Entries("sr").Apply Selection.Range
```

This example applies the "sr" AutoCorrect entry to the first word in the active document.




```
AutoCorrect.Entries("sr").Apply ActiveDocument.Words(1)
```


## See also


#### Concepts


[AutoCorrectEntry Object](autocorrectentry-object-word.md)

