---
title: SmartArtColor Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SmartArtColor
ms.assetid: 5aca0209-20d3-c16f-fdfd-184f3464e00b
---


# SmartArtColor Object (Office)

Chooses the color scheme for the SmartArt diagram.


## Remarks

Simulates the commands on the Microsoft Office Fluent Ribbon user interface on the SmartArt Tools tab, on the Design group, on the Change Colors command.


## Example

The following code sets the color scheme of the Smart Art diagram.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.Color = Application.SmartArtColors(1)
```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

