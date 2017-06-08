---
title: SeriesCollection Object (Excel)
keywords: vbaxl10.chm579072
f1_keywords:
- vbaxl10.chm579072
ms.prod: excel
api_name:
- Excel.SeriesCollection
ms.assetid: 93aa1f0b-4939-8c60-a444-2f791e8ce144
ms.date: 06/08/2017
---


# SeriesCollection Object (Excel)

A collection of all the  **[Series](series-object-excel.md)** objects in the specified chart or chart group.


## Remarks

Use the  **[SeriesCollection](chart-seriescollection-method-excel.md)** method to return the **SeriesCollection** collection.


## Example

 The following example adds the data in cells C1:C10 on worksheet one to an existing series in the series collection in embedded chart one.


```
Worksheets(1).ChartObjects(1).Chart. _ 
 SeriesCollection.Extend Worksheets(1).Range("c1:c10")
```

Use the  **[Add](seriescollection-add-method-excel.md)** method to create a new series and add it to the chart. The following example adds the data from cells A1:A19 as a new series on the chart sheet named "Chart1."




```
Charts("chart1").SeriesCollection.Add _ 
 source:=Worksheets("sheet1").Range("a1:a19")
```

Use  **SeriesCollection** ( _index_ ), where _index_ is the series index number or name, to return a single **Series** object. The following example sets the color of the interior for the first series in embedded chart one on Sheet1.




```
Worksheets("sheet1").ChartObjects(1).Chart. _ 
 SeriesCollection(1).Interior.Color = RGB(255, 0, 0)
```


## Methods



|**Name**|
|:-----|
|[Add](seriescollection-add-method-excel.md)|
|[Extend](seriescollection-extend-method-excel.md)|
|[Item](seriescollection-item-method-excel.md)|
|[NewSeries](seriescollection-newseries-method-excel.md)|
|[Paste](seriescollection-paste-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](seriescollection-application-property-excel.md)|
|[Count](seriescollection-count-property-excel.md)|
|[Creator](seriescollection-creator-property-excel.md)|
|[Parent](seriescollection-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
