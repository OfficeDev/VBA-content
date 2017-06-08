---
title: Application.TwipsToPoints Method (Publisher)
keywords: vbapb10.chm131154
f1_keywords:
- vbapb10.chm131154
ms.prod: publisher
api_name:
- Publisher.Application.TwipsToPoints
ms.assetid: 18e1c4da-1295-31a2-d66b-ab0df807b7a6
ms.date: 06/08/2017
---


# Application.TwipsToPoints Method (Publisher)

Converts a measurement from twips to points (20 twips = 1 point). Returns the converted measurement as a  **Single**.


## Syntax

 _expression_. **TwipsToPoints**( **_Value_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Value|Required| **Single**|The twip value to be converted to points.|

### Return Value

Single


## Remarks

Use the  **[PointsToTwips](application-pointstotwips-method-publisher.md)** method to convert measurements in points to twips.


## Example

This example converts measurements in twips entered by the user to measurements in points.


```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in twips (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) &; " twips = " _ 
 &; Format(Application _ 
 .TwipsToPoints(Value:=Val(strInput)), _ 
 "0.00") &; " points" 
 
 MsgBox strOutput 
Loop 

```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

