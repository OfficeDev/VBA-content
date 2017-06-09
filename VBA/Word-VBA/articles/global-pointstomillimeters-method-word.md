---
title: Global.PointsToMillimeters Method (Word)
keywords: vbawd10.chm163119486
f1_keywords:
- vbawd10.chm163119486
ms.prod: word
api_name:
- Word.Global.PointsToMillimeters
ms.assetid: 0b7c9c70-4352-e427-db1b-4a1b5b2af426
ms.date: 06/08/2017
---


# Global.PointsToMillimeters Method (Word)

Converts a measurement from points to millimeters (1 millimeter = 2.835 points). Returns the converted measurement as a  **Single** .


## Syntax

 _expression_ . **PointsToMillimeters**( **_Points_** )

 _expression_ A variable that represents a **[Global](global-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Points_|Required| **Single**|The measurement, in points.|

### Return Value

Single


## Example

This example converts 72 points to the corresponding number of millimeters.


```vb
MsgBox PointsToMillimeters(72) &; " millimeters"
```

This example converts the value of the variable  _sngData_ (a measurement in points) to centimeters, inches, lines, millimeters, or picas, depending on the value of the variable _intUnit_ (a value from 1 through 5 that indicates the resulting unit of measurement).




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

