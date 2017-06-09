---
title: Application.PointsToPicas Method (Publisher)
keywords: vbapb10.chm131160
f1_keywords:
- vbapb10.chm131160
ms.prod: publisher
api_name:
- Publisher.Application.PointsToPicas
ms.assetid: ff566bef-7032-70f7-7880-ff66cfeca88f
ms.date: 06/08/2017
---


# Application.PointsToPicas Method (Publisher)

Converts a measurement from points to picas (1 pica = 12 points). Returns the converted measurement as a  **Single**.


## Syntax

 _expression_. **PointsToPicas**( **_Value_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Value|Required| **Single**|The point value to be converted to picas.|

### Return Value

Single


## Remarks

Use the  **[PicasToPoints](application-picastopoints-method-publisher.md)** method to convert measurements in picas to points.


## Example

This example converts measurements in points entered by the user to measurements in picas.


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
 .PointsToPicas(Value:=Val(strInput)), _ 
 "0.00") &; " picas" 
 
 MsgBox strOutput 
Loop
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

