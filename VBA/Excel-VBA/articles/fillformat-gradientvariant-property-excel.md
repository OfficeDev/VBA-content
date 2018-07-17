---
title: FillFormat.GradientVariant Property (Excel)
keywords: vbaxl10.chm115016
f1_keywords:
- vbaxl10.chm115016
ms.prod: excel
api_name:
- Excel.FillFormat.GradientVariant
ms.assetid: 00b43056-7d7e-4d5a-edb0-535062fda776
ms.date: 06/08/2017
---


# FillFormat.GradientVariant Property (Excel)

Returns the shade variant for the specified fill as an integer value from 1 through 4. The values for this property correspond to the gradient variants (numbered from left to right and from top to bottom) on the  **Gradient** tab in the **Fill Effects** dialog box. Read-only **Long**


## Syntax

 _expression_ . **GradientVariant**

 _expression_ A variable that represents a **FillFormat** object.


## Remarks

This property is read-only. Use the  **[OneColorGradient](fillformat-onecolorgradient-method-excel.md)** or **[TwoColorGradient](fillformat-twocolorgradient-method-excel.md)** method to set the gradient variant for the fill.


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

