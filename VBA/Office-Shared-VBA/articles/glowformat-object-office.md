---
title: GlowFormat Object (Office)
ms.prod: office
api_name:
- Office.GlowFormat
ms.assetid: b89e2245-e3a4-4a8c-cd4f-86396ad71a5b
ms.date: 06/08/2017
---


# GlowFormat Object (Office)

Represents a glow effect around an Office graphic.


## Example

This example applies glow to the text in the second shape on the second slide in a PowerPoint presentation:


```
With ActivePresentation.Slides(2).Shapes(2) 
 .Text.Font.Glowformat = msoGlowType2 
End With 

```


## Properties



|**Name**|
|:-----|
|[Application](glowformat-application-property-office.md)|
|[Color](glowformat-color-property-office.md)|
|[Creator](glowformat-creator-property-office.md)|
|[Radius](glowformat-radius-property-office.md)|
|[Transparency](glowformat-transparency-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
