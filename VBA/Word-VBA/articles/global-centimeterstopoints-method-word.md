---
title: Global.CentimetersToPoints Method (Word)
keywords: vbawd10.chm163119475
f1_keywords:
- vbawd10.chm163119475
ms.prod: word
api_name:
- Word.Global.CentimetersToPoints
ms.assetid: dc32bb5f-9ea4-e366-d1ad-ac852dc05d82
ms.date: 06/08/2017
---


# Global.CentimetersToPoints Method (Word)

Converts a measurement from centimeters to points (1 cm = 28.35 points). Returns the converted measurement as a  **Single** .


## Syntax

 _expression_ . **CentimetersToPoints**( **_Centimeters_** )

 _expression_ A variable that represents a **[Global](global-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Centimeters_|Required| **Single**|The centimeter value to be converted to points.|

## Example

This example adds a centered tab stop to all the paragraphs in the selection. The tab stop is positioned at 1.5 centimeters from the left margin.


```
Selection.Paragraphs.TabStops.Add _ 
 Position:=CentimetersToPoints(1.5), _ 
 Alignment:=wdAlignTabCenter
```

This example sets a first-line indent of 2.5 centimeters for the first paragraph in the active document.




```vb
ActiveDocument.Paragraphs(1).FirstLineIndent = _ 
 CentimetersToPoints(2.5)
```


## See also


#### Concepts


[Global Object](global-object-word.md)

