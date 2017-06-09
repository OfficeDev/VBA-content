---
title: Application.InchesToPoints Method (Publisher)
keywords: vbapb10.chm131143
f1_keywords:
- vbapb10.chm131143
ms.prod: publisher
api_name:
- Publisher.Application.InchesToPoints
ms.assetid: 32c8740f-ad14-c947-b960-500378a5873d
ms.date: 06/08/2017
---


# Application.InchesToPoints Method (Publisher)

Converts a measurement from inches to points (1 inch = 72 points). Returns the converted measurement as a  **Single**.


## Syntax

 _expression_. **InchesToPoints**( **_Value_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Value|Required| **Single**|The inches value to be converted to points.|

### Return Value

Single


## Remarks

Use the  **[PointsToInches](application-pointstoinches-method-publisher.md)** method to convert measurements in points to inches.


## Example

This example converts measurements in inches entered by the user to measurements in points.


```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in inches (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) &; " in = " _ 
 &; Format(Application _ 
 .InchesToPoints(Value:=Val(strInput)), _ 
 "0.00") &; " points" 
 
 MsgBox strOutput 
Loop 

```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

