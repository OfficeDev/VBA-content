---
title: Charts Object (Excel)
keywords: vbaxl10.chm216072
f1_keywords:
- vbaxl10.chm216072
ms.prod: excel
api_name:
- Excel.Charts
ms.assetid: 06d4602e-a713-7ca0-db39-2d8a29f084a0
ms.date: 06/08/2017
---


# Charts Object (Excel)

A collection of all the chart sheets in the specified or active workbook.


## Remarks

Each chart sheet is represented by a  **Chart** object. This does not include charts embedded on worksheets or dialog sheets. For information about embedded charts, see the **[Chart](chart-object-excel.md)** or **[ChartObject](chartobject-object-excel.md)** topics.


## Example

Use the  **[Charts](workbook-charts-property-excel.md)** property to return the **Charts** collection. The following example prints all chart sheets in the active workbook.


```
Charts.PrintOut
```

Use the  **[Add](http://msdn.microsoft.com/library/370a8ab0-4c65-4a2f-c671-9b5654ff41c0%28Office.15%29.aspx)** method to create a new chart sheet and add it to the workbook. The following example adds a new chart sheet to the active workbook and places the new chart sheet immediately after the worksheet named Sheet1.




```
Charts.Add After:=Worksheets("Sheet1")
```

You can combine the  **Add** method with the **[ChartWizard](chart-chartwizard-method-excel.md)** method to add a new chart that contains data from a worksheet. The following example adds a new line chart based on data in cells A1:A20 on the worksheet named Sheet1.




```
With Charts.Add 
 .ChartWizard source:=Worksheets("Sheet1").Range("A1:A20"), _ 
 Gallery:=xlLine, Title:="February Data" 
End With
```

Use  **Charts** ( _index_ ), where _index_ is the chart-sheet index number or name, to return a single **Chart** object. The following example changes the color of series 1 on chart sheet 1 to red.




```
Charts(1).SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbRed
```

The  **[Sheets](sheets-object-excel.md)** collection contains all the sheets in the workbook (both chart sheets and worksheets). Use **Sheets** ( _index_ ), where _index_ is the sheet name or number, to return a single sheet.


## Methods



|**Name**|
|:-----|
|[Add2](charts-add2-method-excel.md)|
|[Copy](charts-copy-method-excel.md)|
|[Delete](charts-delete-method-excel.md)|
|[Move](charts-move-method-excel.md)|
|[PrintOut](charts-printout-method-excel.md)|
|[PrintPreview](charts-printpreview-method-excel.md)|
|[Select](charts-select-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](charts-application-property-excel.md)|
|[Count](charts-count-property-excel.md)|
|[Creator](charts-creator-property-excel.md)|
|[HPageBreaks](charts-hpagebreaks-property-excel.md)|
|[Item](charts-item-property-excel.md)|
|[Parent](charts-parent-property-excel.md)|
|[Visible](charts-visible-property-excel.md)|
|[VPageBreaks](charts-vpagebreaks-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
