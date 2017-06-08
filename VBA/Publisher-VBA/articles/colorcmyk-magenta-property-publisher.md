---
title: ColorCMYK.Magenta Property (Publisher)
keywords: vbapb10.chm2621444
f1_keywords:
- vbapb10.chm2621444
ms.prod: publisher
api_name:
- Publisher.ColorCMYK.Magenta
ms.assetid: 2996279e-d5f6-9734-ca1a-0e80d7991e5a
ms.date: 06/08/2017
---


# ColorCMYK.Magenta Property (Publisher)

Sets or returns a  **Long** that represents the magenta component of a CMYK color. Value can be any number between 0 and 255. Read/write.


## Syntax

 _expression_. **Magenta**

 _expression_A variable that represents a  **ColorCMYK** object.


### Return Value

Long


## Example

This example creates two new shapes and then sets the CMYK fill color for one shape and sets the CMYK values of the second shape to the same CMYK values.


```vb
Sub ReturnAndSetCMYK() 
 Dim lngCyan As Long 
 Dim lngMagenta As Long 
 Dim lngYellow As Long 
 Dim lngBlack As Long 
 Dim shpHeart As Shape 
 Dim shpStar As Shape 
 
 Set shpHeart = ActiveDocument.Pages(1).Shapes.AddShape _ 
 (Type:=msoShapeHeart, Left:=100, _ 
 Top:=100, Width:=100, Height:=100) 
 Set shpStar = ActiveDocument.Pages(1).Shapes.AddShape _ 
 (Type:=msoShape5pointStar, Left:=200, _ 
 Top:=100, Width:=150, Height:=150) 
 
 With shpHeart.Fill.ForeColor.CMYK 
 .SetCMYK Cyan:=10, Magenta:=80, Yellow:=200, Black:=30 
 lngCyan = .Cyan 
 lngMagenta = .Magenta 
 lngYellow = .Yellow 
 lngBlack = .Black 
 End With 
 
 'Sets new shape to current shapes CMYK colors 
 shpStar.Fill.ForeColor.CMYK.SetCMYK _ 
 Cyan:=lngCyan, Magenta:=lngMagenta, _ 
 Yellow:=lngYellow, Black:=lngBlack 
End Sub
```


