---
title: ChartObjects Object (Excel)
keywords: vbaxl10.chm495072
f1_keywords:
- vbaxl10.chm495072
ms.prod: excel
api_name:
- Excel.ChartObjects
ms.assetid: 67cf2d82-ed9b-b23d-836f-19b106bcc5ed
ms.date: 06/08/2017
---


# ChartObjects Object (Excel)

A collection of all the  **[ChartObject](chartobject-object-excel.md)** objects on the specified chart sheet, dialog sheet, or worksheet.


## Remarks

Each  **ChartObject** object represents an embedded chart. The **ChartObject** object acts as a container for a **[Chart](chart-object-excel.md)** object. Properties and methods for the **ChartObject** object control the appearance and size of the embedded chart on the sheet. **ChartObjects** collection


## Example

Use the  **[ChartObjects](worksheet-chartobjects-method-excel.md)** method to return the **ChartObjects** collection. The following example deletes all the embedded charts on the worksheet named "Sheet1."


```
Worksheets("sheet1").ChartObjects.Delete
```

You cannot use the  **ChartObjects** collection to call the following properties and methods:


-  **Locked** property
    
-  **Placement** property
    
-  **PrintObject** property
    


Unlike in previous version, the  **ChartObjects** collection can now read the properties for height, width, left and top.

Use the  **[Add](chartobjects-add-method-excel.md)** method to create a new, empty embedded chart and add it to the collection. Use the **[ChartWizard](chart-chartwizard-method-excel.md)** method to add data and format the new chart. The following example creates a new embedded chart and then adds the data from cells A1:A20 as a line chart.




```
Dim ch As ChartObject 
Set ch = Worksheets("sheet1").ChartObjects.Add(100, 30, 400, 250) 
ch.Chart.ChartWizard source:=Worksheets("sheet1").Range("a1:a20"), _ 
 gallery:=xlLine, title:="New Chart"
```

Use  **ChartObjects** ( _index_ ), where _index_ is the embedded chart index number or name, to return a single object. The following example sets the pattern for the chart area in embedded Chart 1 on the worksheet named "Sheet1."




```
Worksheets("Sheet1").ChartObjects(1).Chart. _ 
 CChartObjecthartArea.Format.Fill.Pattern = msoPatternLightDownwardDiagonal 
```


## Methods



|**Name**|
|:-----|
|[Add](chartobjects-add-method-excel.md)|
|[Copy](chartobjects-copy-method-excel.md)|
|[CopyPicture](chartobjects-copypicture-method-excel.md)|
|[Cut](chartobjects-cut-method-excel.md)|
|[Delete](chartobjects-delete-method-excel.md)|
|[Duplicate](chartobjects-duplicate-method-excel.md)|
|[Item](chartobjects-item-method-excel.md)|
|[Select](chartobjects-select-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](chartobjects-application-property-excel.md)|
|[Count](chartobjects-count-property-excel.md)|
|[Creator](chartobjects-creator-property-excel.md)|
|[Height](chartobjects-height-property-excel.md)|
|[Left](chartobjects-left-property-excel.md)|
|[Locked](chartobjects-locked-property-excel.md)|
|[Parent](chartobjects-parent-property-excel.md)|
|[Placement](chartobjects-placement-property-excel.md)|
|[PrintObject](chartobjects-printobject-property-excel.md)|
|[ProtectChartObject](chartobjects-protectchartobject-property-excel.md)|
|[ShapeRange](chartobjects-shaperange-property-excel.md)|
|[Top](chartobjects-top-property-excel.md)|
|[Visible](chartobjects-visible-property-excel.md)|
|[Width](chartobjects-width-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
