---
title: WorksheetFunction.Forecast_ETS Method (Excel)
keywords: vbaxl10.chm137468
f1_keywords:
- vbaxl10.chm137468
ms.assetid: de915259-3d2a-485a-8027-290dc9cb95a5
ms.date: 06/08/2017
ms.prod: excel
---


# WorksheetFunction.Forecast_ETS Method (Excel)

Calculates or predicts a future value based on existing (historical) values by using the AAA version of the Exponential Smoothing (ETS) algorithm. 


## Syntax

 _expression_ . **Forecast_ETS**( _Arg1_,  _Arg2_,  _Arg3_,  _Arg4_,  _Arg5_,  _Arg6_)

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|||||
| _Arg1_|Required|DOUBLE|Target Date: the data point for which you want to predict a value. Target date can be date/time or numeric. See Remarks.|
| _Arg2_|Required|VARIANT|Values: the historical values, for which you want to forecast the next points.|
| _Arg3_|Required|VARIANT|Timeline: the independent array or range of dates or numeric data. The values in the timeline must have a consistent step between them and can?t be zero. See Remarks.|
| _Arg4_|Optional|VARIANT|Seasonality: A numeric value. See Remarks.|
| _Arg5_|Optional|VARIANT|Data completions: Although the timeline requires a constant step between data points,  **Forecast_ETS** supports up to 30% missing data, and automatically adjusts for it. See Remarks.|
| _Arg6_|Optional|VARIANT|Aggregation: Although the timeline requires a constant step between data points,  **Forecast_ETS** aggregates multiple points that have the same time stamp. See Remarks.|

### Return Value

 **Double**


## Remarks

The predicted value is a continuation of the historical values in the specified target date, which should be a continuation of the timeline. You can use this function to predict future sales, inventory requirements, or consumer trends.


 **Note**  This function requires the timeline to be organized with a constant step between the different points. For example, that could be a monthly timeline with values on the 1st of every month, a yearly timeline, or a timeline of numerical indices. In general, it?s very useful to aggregate raw detailed data before you apply the forecast, which produces more accurate forecast results as well.

If the target date parameter value ( _Arg1_) is chronologically before the end of the historical timeline,  **Forecast_ETS** returns an error.

 It isn't necessary to sort the timeline ( _Arg3_), because  **Forecast_ETS** sorts it implicitly for calculations. If **Forecast_ETS** can't identify a constant step in the timeline, it returns runtime error ?1004?. If the timeline contains duplicate values, **Forecast_ETS** also returns an error. If the ranges of the timeline and values aren't all of the same size, **Forecast_ETS** returns runtime error ?1004?.

The default value of 1 for seasonality ( _Arg4_) means Excel detects seasonality automatically for the forecast and uses positive, whole numbers for the length of the seasonal pattern. 0 indicates no seasonality, meaning the prediction will be linear. Positive whole numbers indicate to the algorithm to use patterns of this length as the seasonality. For any other value,  **Forecast_ETS** returns an error. Maximum supported seasonality is 8,760 (the number of hours in a year). Any seasonality value above that number results in an error.

Passing 0 for the data completions parameter ( _Arg5_) instructs the algorithm to account for missing points as zeros. The default value of 1 accounts for missing points by computing them to be the average of the neighboring points. If there is more than 30% missing data,  **Forecast_ETS** returns runtime error ?1004?.

The aggregation parameter ( _Arg6_) is a numeric value specifying the method to use to aggregate several values that have the same time stamp. The default value of 0 specifies AVERAGE, while other numbers between 1 and 6 specify SUM, COUNT, COUNTA, MIN, MAX, and MEDIAN.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

