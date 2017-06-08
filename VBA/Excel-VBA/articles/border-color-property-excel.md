---
title: Border.Color Property (Excel)
keywords: vbaxl10.chm547073
f1_keywords:
- vbaxl10.chm547073
ms.prod: excel
api_name:
- Excel.Border.Color
ms.assetid: ca90fc42-2a7a-d43e-9c2c-0055f6bf9010
ms.date: 06/08/2017
---


# Border.Color Property (Excel)

Returns or sets the primary color of the object, as shown in the table in the remarks section. Use the  **RGB** function to create a color value. Read/write **Variant** .


## Syntax

 _expression_ . **Color**

 _expression_ An expression that returns a **Border** object.


## Remarks





|**Object**|**Color**|
|:-----|:-----|
| **Border**|The color of the border.|
| **Borders**|The color of all four borders of a range. If they're not all the same color,  **Color** returns 0 (zero).|
| **Font**|The color of the font.|
| **Interior**|The cell shading color or the drawing object fill color.|
| **Tab**|The color of the tab.|

## Example

This example sets the color of the tick-mark labels on the value axis in Chart1.


```vb
Charts("Chart1").Axes(xlValue).TickLabels.Font.Color = _ 
 RGB(0, 255, 0)
```


## See also


#### Concepts


[Border Object](border-object-excel.md)

