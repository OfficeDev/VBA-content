---
title: FillFormat.TwoColorGradient Method (Publisher)
keywords: vbapb10.chm2359318
f1_keywords:
- vbapb10.chm2359318
ms.prod: publisher
api_name:
- Publisher.FillFormat.TwoColorGradient
ms.assetid: 7b0d1b19-a7bf-7b3d-66f4-60dfc588abfe
ms.date: 06/08/2017
---


# FillFormat.TwoColorGradient Method (Publisher)

Sets the specified fill to a two-color gradient. The two fill colors are specified by the  **[ForeColor](fillformat-forecolor-property-publisher.md)** and **[BackColor](fillformat-backcolor-property-publisher.md)** properties.


## Syntax

 _expression_. **TwoColorGradient**( **_Style_**,  **_Variant_**)

 _expression_A variable that represents a  **FillFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Style|Required| **MsoGradientStyle**|The gradient style.|
|Variant|Required| **Long**|The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the  **Gradient** tab in the **Fill Effects** dialog box. If Style is **msoGradientFromTitle** or **msoGradientFromCenter**, this argument can be either 1 or 2.|

## Remarks

The Style parameter can be one of the  **MsoGradientStyle** constants declared in the Microsoft Office type library and shown in the following table.



| **msoGradientDiagonalDown**|
| **msoGradientDiagonalUp**|
| **msoGradientFromCenter**|
| **msoGradientFromCorner**|
| **msoGradientFromTitle**|
| **msoGradientHorizontal**|
| **msoGradientVertical**|

## Example

This example adds a rectangle with a two-color gradient fill to the active publication and sets the background and foreground color for the fill.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeRectangle, _ 
 Left:=0, Top:=0, Width:=40, Height:=80).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(0, 170, 170) 
 .TwoColorGradient Style:=msoGradientHorizontal, Variant:=1 
End With 

```


