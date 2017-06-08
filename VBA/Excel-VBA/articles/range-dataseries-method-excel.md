---
title: Range.DataSeries Method (Excel)
keywords: vbaxl10.chm144113
f1_keywords:
- vbaxl10.chm144113
ms.prod: excel
api_name:
- Excel.Range.DataSeries
ms.assetid: cfdb0582-8b6c-029d-2a99-4fa1d4b360ea
ms.date: 06/08/2017
---


# Range.DataSeries Method (Excel)

Creates a data series in the specified range.  **Variant** .


## Syntax

 _expression_ . **DataSeries**( **_Rowcol_** , **_Type_** , **_Date_** , **_Step_** , **_Stop_** , **_Trend_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Rowcol_|Optional| **Variant**|Can be the  **xlRows** or **xlColumns** constant to have the data series entered in rows or columns, respectively. If this argument is omitted, the size and shape of the range is used.|
| _Type_|Optional| **[XlDataSeriesType](xldataseriestype-enumeration-excel.md)**|The type for the data series.|
| _Date_|Optional| **[XlDataSeriesDate](xldataseriesdate-enumeration-excel.md)**|If the  _Type_ argument is **xlChronological** , the _Date_ argument indicates the step date unit.|
| _Step_|Optional| **Variant**|The step value for the series. The default value is 1.|
| _Stop_|Optional| **Variant**|The stop value for the series. If this argument is omitted, Microsoft Excel fills to the end of the range.|
| _Trend_|Optional| **Variant**| **True** to create a linear trend or growth trend. **False** to create a standard data series. The default value is **False** .|

### Return Value

Variant


## Example

This example creates a series of 12 dates. The series contains the last day of every month in 1996 and is created in the range A1:A12 on Sheet1.


```vb
Set dateRange = Worksheets("Sheet1").Range("A1:A12") 
Worksheets("Sheet1").Range("A1").Formula = "31-JAN-1996" 
dateRange.DataSeries Type:=xlChronological, Date:=xlMonth
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

