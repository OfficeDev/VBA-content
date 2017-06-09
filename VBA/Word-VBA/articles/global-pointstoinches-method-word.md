---
title: Global.PointsToInches Method (Word)
keywords: vbawd10.chm163119484
f1_keywords:
- vbawd10.chm163119484
ms.prod: word
api_name:
- Word.Global.PointsToInches
ms.assetid: e3d6ab40-3919-55e0-5829-603fca24c226
ms.date: 06/08/2017
---


# Global.PointsToInches Method (Word)

Converts a measurement from points to inches (1 inch = 72 points). Returns the converted measurement as a  **Single** .


## Syntax

 _expression_ . **PointsToInches**( **_Points_** )

 _expression_ A variable that represents a **[Global](global-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Points_|Required| **Single**|The measurement, in points.|

### Return Value

Single


## Example

This example converts the measurement of the top margin for the active document to inches and displays the result in a message box.


```vb
MsgBox PointsToInches(ActiveDocument.Sections(1) _ 
 .PageSetup.TopMargin)
```

This example converts the value of the variable  `sngData` (a measurement in points) to centimeters, inches, lines, millimeters, or picas, depending on the value of the variable `intUnit` (a value from 1 through 5 that indicates the resulting unit of measurement).




```vb
Function ConvertPoints(ByVal intUnit As Integer, _ 
 sngData As Single) As Single 
 
 Select Case intUnit 
 Case 1 
 ConvertPoints = PointsToCentimeters(sngData) 
 Case 2 
 ConvertPoints = PointsToInches(sngData) 
 Case 3 
 ConvertPoints = PointsToLines(sngData) 
 Case 4 
 ConvertPoints = PointsToMillimeters(sngData) 
 Case 5 
 ConvertPoints = PointsToPicas(sngData) 
 Case Else 
 Error 5 
 End Select 
 
End Function
```


## See also


#### Concepts


[Global Object](global-object-word.md)

