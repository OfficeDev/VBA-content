---
title: Application.InchesToPoints Method (Word)
keywords: vbawd10.chm158335346
f1_keywords:
- vbawd10.chm158335346
ms.prod: word
api_name:
- Word.Application.InchesToPoints
ms.assetid: 67a7e59c-bc61-be03-852d-05fadebef148
ms.date: 06/08/2017
---


# Application.InchesToPoints Method (Word)

Converts a measurement from inches to points (1 inch = 72 points). Returns the converted measurement as a  **Single** .


## Syntax

 _expression_ . **InchesToPoints**( **_Inches_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Inches_|Required| **Single**|The inch value to be converted to points.|

## Example

This example sets the space before for the selected paragraphs to 0.25 inch.


```
Selection.ParagraphFormat.SpaceBefore = InchesToPoints(0.25)
```

This example prints each open document after setting the left and right margins to 0.65 inch.




```vb
Dim docLoop As Document 
 
For Each docLoop in Documents 
 With docLoop 
 .PageSetup.LeftMargin = InchesToPoints(0.65) 
 .PageSetup.RightMargin = InchesToPoints(0.65) 
 .PrintOut 
 End With 
Next docLoop
```


## See also


#### Concepts


[Application Object](application-object-word.md)

