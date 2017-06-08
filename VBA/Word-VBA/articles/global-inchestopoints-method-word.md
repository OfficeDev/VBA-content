---
title: Global.InchesToPoints Method (Word)
keywords: vbawd10.chm163119474
f1_keywords:
- vbawd10.chm163119474
ms.prod: word
api_name:
- Word.Global.InchesToPoints
ms.assetid: 7e8f5631-fa6a-702a-5785-da7b34495a22
ms.date: 06/08/2017
---


# Global.InchesToPoints Method (Word)

Converts a measurement from inches to points (1 inch = 72 points). Returns the converted measurement as a  **Single** .


## Syntax

 _expression_ . **InchesToPoints**( **_Inches_** )

 _expression_ A variable that represents a **[Global](global-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Inches_|Required| **Single**|The inch value to be converted to points.|

### Return Value

Single


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


[Global Object](global-object-word.md)

