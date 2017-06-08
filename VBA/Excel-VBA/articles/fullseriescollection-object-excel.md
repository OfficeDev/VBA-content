---
title: FullSeriesCollection Object (Excel)
keywords: vbaxl10.chm943072
f1_keywords:
- vbaxl10.chm943072
ms.prod: excel
ms.assetid: 5d7b7e7c-0a74-307b-84f9-56143ceba464
ms.date: 06/08/2017
---


# FullSeriesCollection Object (Excel)

Represents the full set of [Series Object (Excel)](series-object-excel.md) objects in a chart.


## Remarks

The [FullSeriesCollection Object (Excel)](fullseriescollection-object-excel.md) object enables you to get a filtered out[Series Object (Excel)](series-object-excel.md) object and filter it back in. It also enables you to iterate over the full set of Series object, filtered out or visible, programmatically. By having the existing[SeriesCollection Object (Excel)](seriescollection-object-excel.md) object contain only the visible series you can programmatically perform operations on only the visible series. It also prevents Microsoft Excel from breaking existing chart solutions on charts with filtered out data.


## Example

The following example displays a message box with the name of the second [Series Object (Excel)](series-object-excel.md) object in the second chart.


```VB.net
MsgBox Chart(1).FullSeriesCollection.Item(2).Name
```


## Methods



|**Name**|
|:-----|
|[Item](fullseriescollection-item-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](fullseriescollection-application-property-excel.md)|
|[Count](fullseriescollection-count-property-excel.md)|
|[Creator](fullseriescollection-creator-property-excel.md)|
|[Parent](fullseriescollection-parent-property-excel.md)|

## See also


#### Concepts


[SeriesCollection](seriescollection-object-excel.md)
#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
