---
title: SeriesCollection.Add Method (Excel)
keywords: vbaxl10.chm580074
f1_keywords:
- vbaxl10.chm580074
ms.prod: excel
api_name:
- Excel.SeriesCollection.Add
ms.assetid: 11bedfdb-de8e-94cf-a23d-2c6e1d85cc9a
ms.date: 06/08/2017
---


# SeriesCollection.Add Method (Excel)

Adds one or more new series to the  **SeriesCollection** collection.


## Syntax

 _expression_ . **Add**( **_Source_** , **_Rowcol_** , **_SeriesLabels_** , **_CategoryLabels_** , **_Replace_** )

 _expression_ A variable that represents a **SeriesCollection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Source_|Required| **Variant**|The new data as a  **[Range](range-object-excel.md)** object.|
| _Rowcol_|Optional| **[XlRowCol](xlrowcol-enumeration-excel.md)**|. Specifies whether the new values are in the rows or columns of the specified range.|
| _SeriesLabels_|Optional| **Variant**| **True** if the first row or column contains the name of the data series. **False** if the first row or column contains the first data point of the series. If this argument is omitted, Microsoft Excel attempts to determine the location of the series name from the contents of the first row or column.|
| _CategoryLabels_|Optional| **Variant**| **True** if the first row or column contains the name of the category labels. **False** if the first row or column contains the first data point of the series. If this argument is omitted, Microsoft Excel attempts to determine the location of the category label from the contents of the first row or column.|
| _Replace_|Optional| **Variant**|If  _CategoryLabels_ is **True** and _Replace_ is **True** , the specified categories replace the categories that currently exist for the series. If _Replace_ is **False** , the existing categories will not be replaced. The default value is **False** .|

### Return Value

A  **[Series](series-object-excel.md)** object that represents the new series.


## Remarks

This method does not actually return a  **Series** object as stated in the Object Browser. This method is not available for PivotChart reports.


## Example

This example creates a new series in Chart1. The data source for the new series is range B1:B10 on Sheet1.


```vb
Charts("Chart1").SeriesCollection.Add _ 
 Source:=ActiveWorkbook.Worksheets("Sheet1").Range("B1:B10")
```

This example creates a new series on the embedded chart on Sheet1.




```vb
Worksheets("Sheet1").ChartObjects(1).Activate 
ActiveChart.SeriesCollection.Add _ 
 Source:=Worksheets("Sheet1").Range("B1:B10")
```


## See also


#### Concepts


[SeriesCollection Object](seriescollection-object-excel.md)

