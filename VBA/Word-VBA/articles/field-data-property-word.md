---
title: Field.Data Property (Word)
keywords: vbawd10.chm154075141
f1_keywords:
- vbawd10.chm154075141
ms.prod: word
api_name:
- Word.Field.Data
ms.assetid: b6dfba02-c469-4f8e-e48b-fc69d29673be
ms.date: 06/08/2017
---


# Field.Data Property (Word)

Returns or sets data in an ADDIN field. Read/write  **String** .


## Syntax

 _expression_ . **Data**

 _expression_ A variable that represents a **[Field](field-object-word.md)** object.


## Remarks

The data is not visible in the field code or result; it is only accessible by returning the value of the  **Data** property. If the field isn't an ADDIN field, this property will cause an error.


## Example

This example inserts an ADDIN field at the insertion point in the active document and then sets the data for the field.


```vb
Dim fldTemp As Field 
 
Selection.Collapse Direction:=wdCollapseStart 
Set fldTemp = _ 
 ActiveDocument.Fields.Add(Range:=Selection.Range, _ 
 Type:=wdFieldAddin) 
fldTemp.Data = "Hidden information"
```


## See also


#### Concepts


[Field Object](field-object-word.md)

