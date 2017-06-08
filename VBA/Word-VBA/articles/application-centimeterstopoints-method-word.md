---
title: Application.CentimetersToPoints Method (Word)
keywords: vbawd10.chm158335347
f1_keywords:
- vbawd10.chm158335347
ms.prod: word
api_name:
- Word.Application.CentimetersToPoints
ms.assetid: ca57a957-cc39-49ff-5e51-608e7985fd51
ms.date: 06/08/2017
---


# Application.CentimetersToPoints Method (Word)

Converts a measurement from centimeters to points (1 cm = 28.35 points). Returns the converted measurement as a  **Single** .


## Syntax

 _expression_ . **CentimetersToPoints**( **_Centimeters_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Centimeters_|Required| **Single**|The centimeter value to be converted to points.|

## Example

This example adds a centered tab stop to all the paragraphs in the selection. The tab stop is positioned at 1.5 centimeters from the left margin.


```
Selection.Paragraphs.TabStops.Add _ 
 Position:=Application.CentimetersToPoints(1.5), _ 
 Alignment:=wdAlignTabCenter
```

This example sets a first-line indent of 2.5 centimeters for the first paragraph in the active document.




```vb
ActiveDocument.Paragraphs(1).FirstLineIndent = _ 
 Application.CentimetersToPoints(2.5)
```


## See also


#### Concepts


[Application Object](application-object-word.md)

