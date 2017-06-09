---
title: Application.LinesToPoints Method (Word)
keywords: vbawd10.chm158335350
f1_keywords:
- vbawd10.chm158335350
ms.prod: word
api_name:
- Word.Application.LinesToPoints
ms.assetid: f146db0f-35f6-d25d-2674-e35a7c08801b
ms.date: 06/08/2017
---


# Application.LinesToPoints Method (Word)

Converts a measurement from lines to points (1 line = 12 points). Returns the converted measurement as a  **Single** .


## Syntax

 _expression_ . **LinesToPoints**( **_Lines_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Lines_|Required| **Single**|The line value to be converted to points.|

### Return Value

Single


## Example

This example sets the paragraph line spacing in the selection to three lines.


```vb
With Selection.ParagraphFormat 
 .LineSpacingRule = wdLineSpaceMultiple 
 .LineSpacing = LinesToPoints(3) 
End With
```


## See also


#### Concepts


[Application Object](application-object-word.md)

