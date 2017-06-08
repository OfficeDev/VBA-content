---
title: Application.PixelsToPoints Method (Publisher)
keywords: vbapb10.chm131153
f1_keywords:
- vbapb10.chm131153
ms.prod: publisher
api_name:
- Publisher.Application.PixelsToPoints
ms.assetid: 5d7e453f-e962-e557-48e4-44766d0c64d9
ms.date: 06/08/2017
---


# Application.PixelsToPoints Method (Publisher)

Converts a measurement from pixels to points (1 pixel = 0.75 points). Returns the converted measurement as a  **Single**.


## Syntax

 _expression_. **PixelsToPoints**( **_Value_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Value|Required| **Single**|The pixel value to be converted to points.|

### Return Value

Single


## Remarks

Use the  **[PointsToPixels](application-pointstopixels-method-publisher.md)** method to convert measurements in points to pixels.


## Example

This example converts measurements in pixels entered by the user to measurements in points.


```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in pixels (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) &; " pixels = " _ 
 &; Format(Application _ 
 .PixelsToPoints(Value:=Val(strInput)), _ 
 "0.00") &; " points" 
 
 MsgBox strOutput 
Loop 

```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

