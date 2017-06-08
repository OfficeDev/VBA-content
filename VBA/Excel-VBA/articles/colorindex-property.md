---
title: ColorIndex Property
keywords: vbagr10.chm5207225
f1_keywords:
- vbagr10.chm5207225
ms.prod: excel
api_name:
- Excel.ColorIndex
ms.assetid: e9a9c9de-8a42-0f61-be25-4c158709df68
ms.date: 06/08/2017
---


# ColorIndex Property

Returns or sets the color of the border, font or interior, as shown in the following table. The color is specified as an index value into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Read/write Variant.



|**Object**|**Description**|
|:-----|:-----|
| **Border**|The color of the border.|
| **Font**|The color of the font.|
| **Interior**|The color of the interior fill. Set  **ColorIndex** to **xlColorIndexNone** to specify that you don't want an interior fill. Set **ColorIndex** to **xlColorIndexAutomatic** to specify the automatic fill (for drawing objects).|

 _expression_. **ColorIndex**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Remarks

This property specifies a color as an index into the color palette. The following illustration shows the color-index values in the default color palette.


![Color](images/colorin_ZA06050819.gif)




## Example

The following examples assume that you're using the default color palette.

This example sets the color of the major gridlines for the value axis.




```vb
With myChart.Axes(xlValue) 
 If .HasMajorGridlines Then 
 'Set color to blue 
 .MajorGridlines.Border.ColorIndex = 5 
 End If 
End With
```

This example sets the color of the chart area interior to red and sets the border color to blue.




```vb
With myChart.ChartArea 
 .Interior.ColorIndex = 3 
 .Border.ColorIndex = 5 
End With
```


