---
title: Application.PointsToEmus Method (Publisher)
keywords: vbapb10.chm131156
f1_keywords:
- vbapb10.chm131156
ms.prod: publisher
api_name:
- Publisher.Application.PointsToEmus
ms.assetid: cb3f0bb9-fa0d-d967-9294-081a369c2c4e
ms.date: 06/08/2017
---


# Application.PointsToEmus Method (Publisher)

Converts a measurement from points to emus (12700 emus = 1 point). Returns the converted measurement as a  **Single**.


## Syntax

 _expression_. **PointsToEmus**( **_Value_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Value|Required| **Single**|The point value to be converted to emus.|

### Return Value

Single


## Remarks

Use the  **[EmusToPoints](application-emustopoints-method-publisher.md)** method to convert measurements in emus to points.


## Example

This example converts measurements in points entered by the user to measurements in centimeters.


```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in points (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) &; " points = " _ 
 &; Format(Application _ 
 .PointsToEmus(Value:=Val(strInput)), _ 
 "0.00") &; " emus" 
 
 MsgBox strOutput 
Loop 

```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

