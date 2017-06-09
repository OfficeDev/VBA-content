---
title: Application.CentimetersToPoints Method (Publisher)
keywords: vbapb10.chm131141
f1_keywords:
- vbapb10.chm131141
ms.prod: publisher
api_name:
- Publisher.Application.CentimetersToPoints
ms.assetid: 6eda6692-ea9a-c4ad-6991-066fdc23bd2c
ms.date: 06/08/2017
---


# Application.CentimetersToPoints Method (Publisher)

Converts a measurement from centimeters to points (1 cm = 28.35 points). Returns the converted measurement as a  **Single**.


## Syntax

 _expression_. **CentimetersToPoints**( **_Value_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Value|Required| **Single**|The centimeter value to be converted to points.|

### Return Value

Single


## Remarks

Use the  **[PointsToCentimeters](application-pointstocentimeters-method-publisher.md)** method to convert measurements in points to centimeters.


## Example

This example converts measurements in centimeters entered by the user to measurements in points.


```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in centimeters (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) &; " cm = " _ 
 &; Format(Application _ 
 .CentimetersToPoints(Value:=Val(strInput)), _ 
 "0.00") &; " points" 
 
 MsgBox strOutput 
Loop 

```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

