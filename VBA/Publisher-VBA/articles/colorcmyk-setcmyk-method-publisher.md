---
title: ColorCMYK.SetCMYK Method (Publisher)
keywords: vbapb10.chm2621447
f1_keywords:
- vbapb10.chm2621447
ms.prod: publisher
api_name:
- Publisher.ColorCMYK.SetCMYK
ms.assetid: 9c7ec18b-73e9-66bc-57f4-cd6d62817630
ms.date: 06/08/2017
---


# ColorCMYK.SetCMYK Method (Publisher)

Sets a cyan-magenta-yellow-black (CMYK) color value.


## Syntax

 _expression_. **SetCMYK**( **_Cyan_**,  **_Magenta_**,  **_Yellow_**,  **_Black_**)

 _expression_A variable that represents a  **ColorCMYK** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Cyan|Required| **Long**|A number that represents the cyan component of the color. Value can be any number between 0 and 255.|
|Magenta|Required| **Long**|A number that represents the magenta component of the color. Value can be any number between 0 and 255.|
|Yellow|Required| **Long**|A number that represents the yellow component of the color. Value can be any number between 0 and 255.|
|Black|Required| **Long**|A number that represents the black component of the color. Value can be any number between 0 and 255.|

## Example

This example sets the CMYK color for the specified shape.


```vb
Sub SetCMYKColor() 
 Dim shpStar As Shape 
 
 Set shpStar = ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShape5pointStar, Left:=72, _ 
 Top:=72, Width:=150, Height:=150) 
 shpStar.Fill.ForeColor.CMYK.SetCMYK Cyan:=0, _ 
 Magenta:=255, Yellow:=255, Black:=50 
End Sub
```


