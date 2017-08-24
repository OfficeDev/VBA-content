---
title: ColorCMYK Object (Publisher)
keywords: vbapb10.chm2686975
f1_keywords:
- vbapb10.chm2686975
ms.prod: publisher
api_name:
- Publisher.ColorCMYK
ms.assetid: e1a39f6f-f440-e375-4f8c-e81093e5a451
ms.date: 06/08/2017
---


# ColorCMYK Object (Publisher)

Represents a cyan-magenta-yellow-black (CMYK) color value.
 


## Example

Use the  **CMYK** property of a **ColorFormat** object to return a **ColorCMYK** object. Use the **Cyan**, **Magenta**, **Yellow**, and **Black** properties of the **ColorCMYK** object to individually set each of the four colors in the CMYK color value. Use the **SetCMYK** method on a **ColorCMYK** object to set all four colors at once.
 

 

 

 
The following example retrieves the CMYK color value of shape one's fill and changes it to another CMYK color value.
 

 



```
Dim cmykColor As ColorCMYK Set cmykColor = ActiveDocument.Pages(1).Shapes(1).Fill.ForeColor.CMYK cmykColor.SetCMYK Cyan:=0, Magenta:=255, Yellow:=255, Black:=50
```


## Methods



|**Name**|
|:-----|
|[SetCMYK](colorcmyk-setcmyk-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](colorcmyk-application-property-publisher.md)|
|[Black](colorcmyk-black-property-publisher.md)|
|[Cyan](colorcmyk-cyan-property-publisher.md)|
|[Magenta](colorcmyk-magenta-property-publisher.md)|
|[Parent](colorcmyk-parent-property-publisher.md)|
|[Yellow](colorcmyk-yellow-property-publisher.md)|

