---
title: Field.UpdateSource Method (Word)
keywords: vbawd10.chm154075239
f1_keywords:
- vbawd10.chm154075239
ms.prod: word
api_name:
- Word.Field.UpdateSource
ms.assetid: 8a7a3362-efc5-97e8-c951-e3143e28488d
ms.date: 06/08/2017
---


# Field.UpdateSource Method (Word)

Saves the changes made to the results of an INCLUDETEXT field back to the source document.


## Syntax

 _expression_ . **UpdateSource**

 _expression_ Required. A variable that represents a **[Field](field-object-word.md)** object.


## Remarks

The source document must be formatted as a Word document.


## Example

This example updates the INCLUDETEXT fields in the active document.


```vb
Dim fldLoop As Field 
 
For Each fldLoop In ActiveDocument.Fields 
 If fldLoop.Type = wdFieldIncludeText Then _ 
 fldLoop.UpdateSource 
Next fldLoop
```


## See also


#### Concepts


[Field Object](field-object-word.md)

