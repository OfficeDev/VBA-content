---
title: Axis Object (Excel)
keywords: vbaxl10.chm560072
f1_keywords:
- vbaxl10.chm560072
ms.prod: excel
api_name:
- Excel.Axis
ms.assetid: 7e08c61b-90f4-8d91-0ee2-84283d10b324
ms.date: 06/08/2017
---


# Axis Object (Excel)

Represents a single axis in a chart.


## Remarks

The  **Axis** object is a member of the **[Axes](axes-object-excel.md)** collection.

Use  **Axes** ( _type_, _group_ ) where _type_ is the axis type and _group_ is the axis group to return a single **Axis** object. _Type_ can be one of the following **[XlAxisType](xlaxistype-enumeration-excel.md)** constants: **xlCategory**, **xlSeries**, or **xlValue**. _Group_ can be one of the following **[XlAxisGroup](xlaxisgroup-enumeration-excel.md)** constants: **xlPrimary** or **xlSecondary**. For more information, see the **[Axes](chart-axes-method-excel.md)** method.


## Example

The following example sets the category axis title text on the chart sheet named "Chart1."


```
With Charts("chart1").Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Caption = "1994" 
End With
```


## Methods



|**Name**|
|:-----|
|[Delete](axis-delete-method-excel.md)|
|[Select](axis-select-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](axis-application-property-excel.md)|
|[AxisBetweenCategories](axis-axisbetweencategories-property-excel.md)|
|[AxisGroup](axis-axisgroup-property-excel.md)|
|[AxisTitle](axis-axistitle-property-excel.md)|
|[BaseUnit](axis-baseunit-property-excel.md)|
|[BaseUnitIsAuto](axis-baseunitisauto-property-excel.md)|
|[Border](axis-border-property-excel.md)|
|[CategoryNames](axis-categorynames-property-excel.md)|
|[CategoryType](axis-categorytype-property-excel.md)|
|[Creator](axis-creator-property-excel.md)|
|[Crosses](axis-crosses-property-excel.md)|
|[CrossesAt](axis-crossesat-property-excel.md)|
|[DisplayUnit](axis-displayunit-property-excel.md)|
|[DisplayUnitCustom](axis-displayunitcustom-property-excel.md)|
|[DisplayUnitLabel](axis-displayunitlabel-property-excel.md)|
|[Format](axis-format-property-excel.md)|
|[HasDisplayUnitLabel](axis-hasdisplayunitlabel-property-excel.md)|
|[HasMajorGridlines](axis-hasmajorgridlines-property-excel.md)|
|[HasMinorGridlines](axis-hasminorgridlines-property-excel.md)|
|[HasTitle](axis-hastitle-property-excel.md)|
|[Height](axis-height-property-excel.md)|
|[Left](axis-left-property-excel.md)|
|[LogBase](axis-logbase-property-excel.md)|
|[MajorGridlines](axis-majorgridlines-property-excel.md)|
|[MajorTickMark](axis-majortickmark-property-excel.md)|
|[MajorUnit](axis-majorunit-property-excel.md)|
|[MajorUnitIsAuto](axis-majorunitisauto-property-excel.md)|
|[MajorUnitScale](axis-majorunitscale-property-excel.md)|
|[MaximumScale](axis-maximumscale-property-excel.md)|
|[MaximumScaleIsAuto](axis-maximumscaleisauto-property-excel.md)|
|[MinimumScale](axis-minimumscale-property-excel.md)|
|[MinimumScaleIsAuto](axis-minimumscaleisauto-property-excel.md)|
|[MinorGridlines](axis-minorgridlines-property-excel.md)|
|[MinorTickMark](axis-minortickmark-property-excel.md)|
|[MinorUnit](axis-minorunit-property-excel.md)|
|[MinorUnitIsAuto](axis-minorunitisauto-property-excel.md)|
|[MinorUnitScale](axis-minorunitscale-property-excel.md)|
|[Parent](axis-parent-property-excel.md)|
|[ReversePlotOrder](axis-reverseplotorder-property-excel.md)|
|[ScaleType](axis-scaletype-property-excel.md)|
|[TickLabelPosition](axis-ticklabelposition-property-excel.md)|
|[TickLabels](axis-ticklabels-property-excel.md)|
|[TickLabelSpacing](axis-ticklabelspacing-property-excel.md)|
|[TickLabelSpacingIsAuto](axis-ticklabelspacingisauto-property-excel.md)|
|[TickMarkSpacing](axis-tickmarkspacing-property-excel.md)|
|[Top](axis-top-property-excel.md)|
|[Type](axis-type-property-excel.md)|
|[Width](axis-width-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
