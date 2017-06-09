---
title: Application.MillimetersToPoints Method (Word)
keywords: vbawd10.chm158335348
f1_keywords:
- vbawd10.chm158335348
ms.prod: word
api_name:
- Word.Application.MillimetersToPoints
ms.assetid: 13cf2786-709a-d473-0b6d-4fddabb465b8
ms.date: 06/08/2017
---


# Application.MillimetersToPoints Method (Word)

Converts a measurement from millimeters to points (1 mm = 2.85 points). Returns the converted measurement as a  **Single** .


## Syntax

 _expression_ . **MillimetersToPoints**( **_Millimeters_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Millimeters_|Required| **Single**|The millimeter value to be converted to points.|

### Return Value

Single


## Example

This example sets the hyphenation zone in the active document to 8.8 millimeters.


```vb
ActiveDocument.HyphenationZone = MillimetersToPoints(8.8)
```

This example expands the spacing of the selected characters to 2.8 points.




```
Selection.Font.Spacing = MillimetersToPoints(1)
```


## See also


#### Concepts


[Application Object](application-object-word.md)

