---
title: Application.MillimetersToPoints Method (Publisher)
keywords: vbapb10.chm131145
f1_keywords:
- vbapb10.chm131145
ms.prod: publisher
api_name:
- Publisher.Application.MillimetersToPoints
ms.assetid: 40ec9abd-cc1e-9f44-3312-d6689b4822e4
ms.date: 06/08/2017
---


# Application.MillimetersToPoints Method (Publisher)

Converts a measurement from millimeters to points (1 mm = 2.835 points). Returns the converted measurement as a  **Single**.


## Syntax

 _expression_. **MillimetersToPoints**( **_Value_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Value|Required| **Single**|The millimeter value to be converted to points.|

### Return Value

Single


## Remarks

Use the  **[PointsToMillimeters](application-pointstomillimeters-method-publisher.md)** method to convert measurements in points to millimeters.


## Example

This example converts measurements in millimeters entered by the user to measurements in points.


```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in millimeters (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) &; " mm = " _ 
 &; Format(Application _ 
 .Mill imetersToPoints(Value:=Val(strInput)), _ 
 "0.00") &; " points" 
 
 MsgBox strOutput 
Loop
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

