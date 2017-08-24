---
title: Application.PointsToTwips Method (Publisher)
keywords: vbapb10.chm131168
f1_keywords:
- vbapb10.chm131168
ms.prod: publisher
api_name:
- Publisher.Application.PointsToTwips
ms.assetid: ba928b83-f551-049e-5868-098a9837ee7b
ms.date: 06/08/2017
---


# Application.PointsToTwips Method (Publisher)

Converts a measurement from points to twips (20 twips = 1 point). Returns the converted measurement as a  **Single**.


## Syntax

 _expression_. **PointsToTwips**( **_Value_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Value|Required| **Single**|The point value to be converted to twips.|

### Return Value

Single


## Remarks

Use the  **[TwipsToPoints](application-twipstopoints-method-publisher.md)** method to convert measurements in twips to points.


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
 .PointsToTwips(Value:=Val(strInput)), _ 
 "0.00") &; " twips" 
 
 MsgBox strOutput 
Loop 

```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

