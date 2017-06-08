---
title: Application.PointsToLines Method (Publisher)
keywords: vbapb10.chm131158
f1_keywords:
- vbapb10.chm131158
ms.prod: publisher
api_name:
- Publisher.Application.PointsToLines
ms.assetid: beab39fe-9458-6878-ae45-487a8b2271df
ms.date: 06/08/2017
---


# Application.PointsToLines Method (Publisher)

Converts a measurement from points to lines (1 line = 12 points). Returns the converted measurement as a  **Single**.


## Syntax

 _expression_. **PointsToLines**( **_Value_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Value|Required| **Single**|The point value to be converted to lines.|

### Return Value

Single


## Remarks

This method assumes a measurement in 12-point lines — the actual size of any text in the publication has no effect on the conversion factor.

Use the  **[LinesToPoints](application-linestopoints-method-publisher.md)** method to convert measurements in lines to points.


## Example

This example converts measurements in lines to measurements in points, demonstrating that the font size in the current selection has no bearing on the conversion factor. Some text must be selected in the active publication for this example to work.


```vb
Dim strOutput As String 
 
' Set text size to 10 points. 
Selection.TextRange.Font.Size = 10 
 
' Display result for 12 points. 
strOutput = "12 points = " _ 
 &; Format(Application _ 
 .PointsToLines(Value:=12), _ 
 "0.00") &; " lines"
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

