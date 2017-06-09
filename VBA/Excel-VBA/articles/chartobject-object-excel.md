---
title: ChartObject Object (Excel)
keywords: vbaxl10.chm493072
f1_keywords:
- vbaxl10.chm493072
ms.prod: excel
api_name:
- Excel.ChartObject
ms.assetid: b546e6f2-7ac6-2dea-eba2-f98f68f3df65
ms.date: 06/08/2017
---


# ChartObject Object (Excel)

Represents an embedded chart on a worksheet.


## Remarks

The  **ChartObject** object acts as a container for a **[Chart](chart-object-excel.md)** object. Properties and methods for the **ChartObject** object control the appearance and size of the embedded chart on the worksheet. The **ChartObject** object is a member of the **[ChartObjects](chartobjects-object-excel.md)** collection. The **ChartObjects** collection contains all the embedded charts on a single sheet.

Use  **ChartObjects** ( _index_ ), where _index_ is the embedded chart index number or name, to return a single **ChartObject** object.


## Example

The following example sets the pattern for the chart area in embedded Chart 1 on the worksheet named "Sheet1."


```
Worksheets("Sheet1").ChartObjects(1).Chart. _ 
 ChartArea.Format.Fill.Pattern = msoPatternLightDownwardDiagonal
```

The embedded chart name is shown in the Name box when the embedded chart is selected. Use the  **[Name](chartobject-name-property-excel.md)** property to set or return the name of the **ChartObject** object. The following example puts rounded corners on the embedded chart named "Chart 1" on the worksheet named "Sheet1."




```
Worksheets("sheet1").ChartObjects("chart 1").RoundedCorners = True
```


## Methods



|**Name**|
|:-----|
|[Activate](chartobject-activate-method-excel.md)|
|[BringToFront](chartobject-bringtofront-method-excel.md)|
|[Copy](chartobject-copy-method-excel.md)|
|[CopyPicture](chartobject-copypicture-method-excel.md)|
|[Cut](chartobject-cut-method-excel.md)|
|[Delete](chartobject-delete-method-excel.md)|
|[Duplicate](chartobject-duplicate-method-excel.md)|
|[Select](chartobject-select-method-excel.md)|
|[SendToBack](chartobject-sendtoback-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](chartobject-application-property-excel.md)|
|[BottomRightCell](chartobject-bottomrightcell-property-excel.md)|
|[Chart](chartobject-chart-property-excel.md)|
|[Creator](chartobject-creator-property-excel.md)|
|[Height](chartobject-height-property-excel.md)|
|[Index](chartobject-index-property-excel.md)|
|[Left](chartobject-left-property-excel.md)|
|[Locked](chartobject-locked-property-excel.md)|
|[Name](chartobject-name-property-excel.md)|
|[Parent](chartobject-parent-property-excel.md)|
|[Placement](chartobject-placement-property-excel.md)|
|[PrintObject](chartobject-printobject-property-excel.md)|
|[ProtectChartObject](chartobject-protectchartobject-property-excel.md)|
|[RoundedCorners](chartobject-roundedcorners-property-excel.md)|
|[Shadow](chartobject-shadow-property-excel.md)|
|[ShapeRange](chartobject-shaperange-property-excel.md)|
|[Top](chartobject-top-property-excel.md)|
|[TopLeftCell](chartobject-topleftcell-property-excel.md)|
|[Visible](chartobject-visible-property-excel.md)|
|[Width](chartobject-width-property-excel.md)|
|[ZOrder](chartobject-zorder-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
