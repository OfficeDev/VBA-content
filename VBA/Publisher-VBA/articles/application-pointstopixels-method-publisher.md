---
title: Application.PointsToPixels Method (Publisher)
keywords: vbapb10.chm131161
f1_keywords:
- vbapb10.chm131161
ms.prod: publisher
api_name:
- Publisher.Application.PointsToPixels
ms.assetid: 9c67fcae-6c93-ddae-cbad-75356e5c5084
ms.date: 06/08/2017
---


# Application.PointsToPixels Method (Publisher)

Converts a measurement from points to pixels (1 pixel = 0.75 points). Returns the converted measurement as a  **Single**.


## Syntax

 _expression_. **PointsToPixels**( **_Value_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Value|Required| **Single**|The point value to be converted to pixels.|

### Return Value

Single


## Remarks

Use the  **[PixelsToPoints](application-pixelstopoints-method-publisher.md)** method to convert measurements in pixels to points.


## Example

This example converts measurements in points entered by the user to measurements in pixels.


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
 .PointsToPixels(Value:=Val(strInput)), _ 
 "0.00") &; " pixels" 
 
 MsgBox strOutput 
Loop 

```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

