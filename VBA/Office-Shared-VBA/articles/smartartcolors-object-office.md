---
title: SmartArtColors Object (Office)
ms.prod: office
api_name:
- Office.SmartArtColors
ms.assetid: a1929517-b1fb-c6fe-b6db-03f7ef1ef894
ms.date: 06/08/2017
---


# SmartArtColors Object (Office)

A collection of SmartArt color styles.


## Remarks

Simulates the commands on the Microsoft Office Fluent Ribbon user interface on the SmartArt Tools, on the Design group, on the Change Colors command.


## Example

The following code sets the color scheme of the Smart Art diagram.


```
ActivePresentation.Slides(1).Shapes(1).SmartArt.Color = Application.SmartArtColors(1)
```


## Methods



|**Name**|
|:-----|
|[Item](smartartcolors-item-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](smartartcolors-application-property-office.md)|
|[Count](smartartcolors-count-property-office.md)|
|[Creator](smartartcolors-creator-property-office.md)|
|[Parent](smartartcolors-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
