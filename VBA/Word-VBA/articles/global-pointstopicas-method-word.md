---
title: Global.PointsToPicas Method (Word)
keywords: vbawd10.chm163119487
f1_keywords:
- vbawd10.chm163119487
ms.prod: word
api_name:
- Word.Global.PointsToPicas
ms.assetid: 7fea77c5-0cc8-ca5e-636b-37400493a6e0
ms.date: 06/08/2017
---


# Global.PointsToPicas Method (Word)

Converts a measurement from points to picas (1 pica = 12 points). Returns the converted measurement as a  **Single** .


## Syntax

 _expression_ . **PointsToPicas**( **_Points_** )

 _expression_ A variable that represents a **[Global](global-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Points_|Required| **Single**|The measurement, in points.|

### Return Value

Single


## Example

This example converts 36 points to the corresponding number of picas.


```vb
MsgBox PointsToPicas(36) &; " picas"
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

