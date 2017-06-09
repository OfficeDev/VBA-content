---
title: ColorScaleCriteria Object (Excel)
keywords: vbaxl10.chm807072
f1_keywords:
- vbaxl10.chm807072
ms.prod: excel
api_name:
- Excel.ColorScaleCriteria
ms.assetid: 9c50a2e4-aa22-92ca-6cef-2f8fc931ec33
ms.date: 06/08/2017
---


# ColorScaleCriteria Object (Excel)

A collection of  **[ColorScaleCriterion](colorscalecriterion-object-excel.md)** objects that represents all of the criteria for a color scale conditional format. Each criterion specifies the minimum, midpoint, or maximum threshold for the color scale.


## Remarks

To return the  **ColorScaleCriteria** collection, use the **[ColorScaleCriteria](colorscale-colorscalecriteria-property-excel.md)** property of the **[ColorScale](colorscale-object-excel.md)** object.


## Example

The following code example creates a range of numbers and then applies a two-color scale conditional formatting rule to that range. The color for the minimum threshold is then assigned to red and the maximum threshold to blue by indexing into the  **ColorScaleCriteria** collection to set individual criteria.


```vb
Sub CreateColorScaleCF() 
 
 Dim cfColorScale As ColorScale 
 
 'Fill cells with sample data from 1 to 10 
 With ActiveSheet 
 .Range("C1") = 1 
 .Range("C2") = 2 
 .Range("C1:C2").AutoFill Destination:=Range("C1:C10") 
 End With 
 
 Range("C1:C10").Select 
 
 'Create a two-color ColorScale object for the created sample data range 
 Set cfColorScale = Selection.FormatConditions.AddColorScale(ColorScaleType:=2) 
 
 'Set the minimum threshold to red and maximum threshold to blue 
 cfColorScale.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0) 
 cfColorScale.ColorScaleCriteria(2).FormatColor.Color = RGB(0, 0, 255) 
 
End Sub
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


