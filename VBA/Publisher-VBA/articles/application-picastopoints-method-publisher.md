---
title: Application.PicasToPoints Method (Publisher)
keywords: vbapb10.chm131152
f1_keywords:
- vbapb10.chm131152
ms.prod: publisher
api_name:
- Publisher.Application.PicasToPoints
ms.assetid: 64d3e435-dcc1-d637-7aac-cc9a9bf81e76
ms.date: 06/08/2017
---


# Application.PicasToPoints Method (Publisher)

Converts a measurement from picas to points (1 pica = 12 points). Returns the converted measurement as a  **Single**.


## Syntax

 _expression_. **PicasToPoints**( **_Value_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Value|Required| **Single**|The pica value to be converted to points.|

### Return Value

Single


## Remarks

Use the  **[PointsToPicas](application-pointstopicas-method-publisher.md)** method to convert measurements in points to picas.


## Example

This example converts measurements in picas entered by the user to measurements in points.


```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in picas (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) &; " picas = " _ 
 &; Format(Application _ 
 .Picas ToPoints(Value:=Val(strInput)), _ 
 "0.00") &; " points" 
 
 MsgBox strOutput 
Loop
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

