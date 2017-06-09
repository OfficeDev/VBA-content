---
title: Application.LinesToPoints Method (Publisher)
keywords: vbapb10.chm131144
f1_keywords:
- vbapb10.chm131144
ms.prod: publisher
api_name:
- Publisher.Application.LinesToPoints
ms.assetid: 55c531aa-5619-6f7f-54e7-7721cb70640e
ms.date: 06/08/2017
---


# Application.LinesToPoints Method (Publisher)

Converts a measurement from lines to points (1 line = 12 points). Returns the converted measurement as a  **Single**.


## Syntax

 _expression_. **LinesToPoints**( **_Value_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Value|Required| **Single**|The line value to be converted to points.|

### Return Value

Single


## Remarks

This method assumes a measurement in 12-point lines — the actual size of any text in the publication has no effect on the conversion factor.

Use the  **[PointsToLines](application-pointstolines-method-publisher.md)** method to convert measurements in points to lines.


## Example

This example converts measurements in lines to measurements in points, demonstrating that the font size in the current selection has no bearing on the conversion factor. Some text must be selected in the active publication for this example to work.


```vb
Dim strOutput As String 
 
' Set text size to 10 points. 
Selection.TextRange.Font.Size = 10 
 
' Display result for one line of text. 
strOutput = "1 line = " _ 
 &; Format(Application _ 
 .LinesToPoints(Value:=1), _ 
 "0.00") &; " points"
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

