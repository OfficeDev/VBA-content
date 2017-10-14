---
title: FillFormat.GradientDegree Property (Excel)
keywords: vbaxl10.chm115014
f1_keywords:
- vbaxl10.chm115014
ms.prod: excel
api_name:
- Excel.FillFormat.GradientDegree
ms.assetid: 46529845-6ee0-def7-8dac-bb459d6ed2f0
ms.date: 06/08/2017
---


# FillFormat.GradientDegree Property (Excel)

Returns the gradient degree of the specified one-color shaded fill as a floating-point value from 0.0 (dark) through 1.0 (light). Read-only  **Single** .


## Syntax

 _expression_ . **GradientDegree**

 _expression_ A variable that represents a **FillFormat** object.


## Remarks

This property is read-only. Use the  **[OneColorGradient](fillformat-onecolorgradient-method-excel.md)** method to set the gradient degree for the fill.


## Example

This example sets the fill format for chart two to the same style used for chart one.


```vb
Set c1f = Charts(1).ChartArea.Fill 
If c1f.Type = msoFillGradient And _ 
 c1f.GradientColorType = msoGradientOneColor Then 
 With Charts(2).ChartArea.Fill 
 .Visible = True 
 .OneColorGradient c1f.GradientStyle, _ 
 c1f.GradientVariant, c1f.GradientDegree 
 End With 
End If
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-excel.md)

