---
title: Interior Object (Excel)
keywords: vbaxl10.chm550072
f1_keywords:
- vbaxl10.chm550072
ms.prod: excel
api_name:
- Excel.Interior
ms.assetid: 37c79831-2cac-69fd-10ee-6d5415ed338b
ms.date: 06/08/2017
---


# Interior Object (Excel)

Represents the interior of an object.


## Example

Use the  **[Interior](range-interior-property-excel.md)** property to return the **Interior** object. The following example sets the color for the interior of cell A1 to red.


```
Worksheets("Sheet1").Range("A1").Interior.ColorIndex = 3
```

 **Sample code provided by:** Holy Macro! Books,[Holy Macro! It's 2,500 Excel VBA Examples](http://www.mrexcel.com/store/index.php?l=product_detail&amp;p=1)

This example gets the value of the color of a cell in column A using the  **ColorIndex** property, and then uses that value to sort the range by color.




```
Sub ColorSort()
   'Set up your variables and turn off screen updating.
   Dim iCounter As Integer
   Application.ScreenUpdating = False
   
   'For each cell in column A, go through and place the color index value of the cell in column C.
   For iCounter = 2 To 55
      Cells(iCounter, 3) = _
         Cells(iCounter, 1).Interior.ColorIndex
   Next iCounter
   
   'Sort the rows based on the data in column C
   Range("C1") = "Index"
   Columns("A:C").Sort key1:=Range("C2"), _
      order1:=xlAscending, header:=xlYes
   
   'Clear out the temporary sorting value in column C, and turn screen updating back on.
   Columns(3).ClearContents
   Application.ScreenUpdating = True
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


## Properties
<a name="AboutContributor"> </a>



|**Name**|
|:-----|
|[Application](interior-application-property-excel.md)|
|[Color](interior-color-property-excel.md)|
|[ColorIndex](interior-colorindex-property-excel.md)|
|[Creator](interior-creator-property-excel.md)|
|[Gradient](interior-gradient-property-excel.md)|
|[InvertIfNegative](interior-invertifnegative-property-excel.md)|
|[Parent](interior-parent-property-excel.md)|
|[Pattern](interior-pattern-property-excel.md)|
|[PatternColor](interior-patterncolor-property-excel.md)|
|[PatternColorIndex](interior-patterncolorindex-property-excel.md)|
|[PatternThemeColor](interior-patternthemecolor-property-excel.md)|
|[PatternTintAndShade](interior-patterntintandshade-property-excel.md)|
|[ThemeColor](interior-themecolor-property-excel.md)|
|[TintAndShade](interior-tintandshade-property-excel.md)|

## See also
<a name="AboutContributor"> </a>


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
