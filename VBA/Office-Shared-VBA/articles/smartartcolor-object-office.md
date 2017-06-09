---
title: SmartArtColor Object (Office)
ms.prod: office
api_name:
- Office.SmartArtColor
ms.assetid: 5aca0209-20d3-c16f-fdfd-184f3464e00b
ms.date: 06/08/2017
---


# SmartArtColor Object (Office)

Chooses the color scheme for the SmartArt diagram.


## Remarks

Simulates the commands on the Microsoft Office Fluent Ribbon user interface on the SmartArt Tools tab, on the Design group, on the Change Colors command.


## Example

The following code sets the color scheme of the Smart Art diagram.


```
ActivePresentation.Slides(1).Shapes(1).SmartArt.Color = Application.SmartArtColors(1)
```


## Properties



|**Name**|
|:-----|
|[Application](smartartcolor-application-property-office.md)|
|[Category](smartartcolor-category-property-office.md)|
|[Creator](smartartcolor-creator-property-office.md)|
|[Description](smartartcolor-description-property-office.md)|
|[Id](smartartcolor-id-property-office.md)|
|[Name](smartartcolor-name-property-office.md)|
|[Parent](smartartcolor-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
