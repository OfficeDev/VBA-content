---
title: WorksheetFunction.Forecast_ETS_Seasonality Method (Excel)
keywords: vbaxl10.chm137470
f1_keywords:
- vbaxl10.chm137470
ms.assetid: aad7c233-1745-64e3-22a9-ade62e5e177d
ms.date: 06/08/2017
ms.prod: excel
---


# WorksheetFunction.Forecast_ETS_Seasonality Method (Excel)

Returns the length of the repetitive pattern Excel detects for the specified time series.


## Syntax

 _expression_ . **Forecast_ETS_Seasonality**( _Arg1_,  _Arg1_,  _Arg2_,  _Arg3_,  _Arg4_)

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|||||
| _Arg1_|Required|VARIANT|Values: the historical values, for which you want to forecast the next points.|
| _Arg2_|Required|VARIANT|Timeline: the independent array or range of dates or numeric data. The values in the timeline must have a consistent step between them and can?t be zero. See Remarks.|
| _Arg3_|Optional|VARIANT|Data completions: Although the timeline requires a constant step between data points,  **Forecast_ETS_Seasonality** supports up to 30% missing data, and automatically adjusts for it. See Remarks.|
| _Arg4_|Optional|VARIANT|Aggregation: Although the timeline requires a constant step between data points,  **Forecast_ETS_Seasonality** aggregates multiple points that have the same time stamp. See Remarks.|

### Return Value

 **Double**


## Remarks

You can use  **Forecast_ETS_Seasonality** following[Forecast_ETS](worksheetfunction-forecast_ets-method-excel.md) to identify which automatic seasonality was detected and used in **Forecast_ETS** . While you can also use it independently of **Forecast_ETS** , the methods are tied together, because the seasonality detected in this method is identical to the one used by **Forecast_ETS** , considering that the same input parameters that affect data completion are passed in both methods.

 It isn't necessary to sort the timeline ( _Arg2_), because  **Forecast_ETS_Seasonality** sorts it implicitly for calculations. If **Forecast_ETS_Seasonality** can't identify a constant step in the timeline, it returns runtime error ?1004?. If the timeline contains duplicate values, **Forecast_ETS_Seasonality** also returns an error. If the ranges of the timeline and values aren't all of the same size, **Forecast_ETS_Seasonality** returns runtime error ?1004?.

Passing 0 for the data completions parameter ( _Arg3_) instructs the algorithm to account for missing points as zeros. The default value of 1 accounts for missing points by computing them to be the average of the neighboring points. If there is more than 30% missing data,  **Forecast_ETS_Seasonality** returns runtime error ?1004?.

The aggregation parameter ( _Arg4_) is a numeric value specifying the method to use to aggregate several values that have the same time stamp. The default value of 0 specifies AVERAGE, while other numbers between 1 and 6 specify SUM, COUNT, COUNTA, MIN, MAX, and MEDIAN.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

