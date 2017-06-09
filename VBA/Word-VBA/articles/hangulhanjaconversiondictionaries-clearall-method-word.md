---
title: HangulHanjaConversionDictionaries.ClearAll Method (Word)
keywords: vbawd10.chm165675110
f1_keywords:
- vbawd10.chm165675110
ms.prod: word
api_name:
- Word.HangulHanjaConversionDictionaries.ClearAll
ms.assetid: 920a8b08-0475-131a-28cc-58cbeb8b6a9c
ms.date: 06/08/2017
---


# HangulHanjaConversionDictionaries.ClearAll Method (Word)

Unloads all of the custom or conversion dictionaries.


## Syntax

 _expression_ . **ClearAll**

 _expression_ Required. A variable that represents a **[HangulHanjaConversionDictionaries](hangulhanjaconversiondictionaries-object-word.md)** collection.


## Remarks

The  **ClearAll** method does not delete the conversion dictionary files. After using this method, the number of conversion dictionaries in the collection is 0 (zero).


## Example

This example unloads all of the conversion dictionaries.


```
HangulHanjaDictionaries.ClearAll
```


## See also


#### Concepts


[HangulHanjaConversionDictionaries Collection Object](hangulhanjaconversiondictionaries-object-word.md)

