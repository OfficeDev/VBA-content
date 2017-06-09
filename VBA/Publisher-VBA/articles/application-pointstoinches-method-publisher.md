---
title: Application.PointsToInches Method (Publisher)
keywords: vbapb10.chm131157
f1_keywords:
- vbapb10.chm131157
ms.prod: publisher
api_name:
- Publisher.Application.PointsToInches
ms.assetid: 58bfd9ce-dee7-0a14-8ec1-7e16a5e967d8
ms.date: 06/08/2017
---


# Application.PointsToInches Method (Publisher)

Converts a measurement from points to inches (1 in = 72 points). Returns the converted measurement as a  **Single**.


## Syntax

 _expression_. **PointsToInches**( **_Value_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Value|Required| **Single**|The point value to be converted to inches.|

### Return Value

Single


## Remarks

Use the  **[InchesToPoints](application-inchestopoints-method-publisher.md)** method to convert measurements in inches to points.


## Example

This example converts measurements in points entered by the user to measurements in inches.


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
 .PointsToInches(Value:=Val(strInput)), _ 
 "0.00") &; " in" 
 
 MsgBox strOutput 
Loop 

```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

