---
title: WorksheetFunction.Forecast_ETS_ConfInt Method (Excel)
keywords: vbaxl10.chm137469
f1_keywords:
- vbaxl10.chm137469
ms.assetid: 23d6cb35-58c8-6ef0-ed4f-5c693974ccd2
ms.date: 06/08/2017
ms.prod: excel
---


# WorksheetFunction.Forecast_ETS_ConfInt Method (Excel)

Returns a confidence interval for the forecast value at the specified target date.


## Syntax

 _expression_ . **Forecast_ETS_ConfInt**( _Arg1_,  _Arg2_,  _Arg3_,  _Arg4_,  _Arg5_,  _Arg6_,  _Arg7_)

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|||||
| _Arg1_|Required|DOUBLE|Target Date: the data point for which you want to predict a value. Target date can be date/time or numeric. See Remarks.|
| _Arg2_|Required|VARIANT|Values: the historical values, for which you want to forecast the next points.|
| _Arg3_|Required|VARIANT|Timeline: the independent array or range of dates or numeric data. The values in the timeline must have a consistent step between them and can?t be zero. See Remarks.|
| _Arg4_|Optional|VARIANT|Confidence level: A numerical value between 0 and 1 (exclusive), indicating a confidence level for the calculated confidence interval. See Remarks.|
| _Arg5_|Optional|VARIANT|Seasonality: A numeric value. See Remarks.|
| _Arg6_|Optional|VARIANT|Data completions: Although the timeline requires a constant step between data points,  **Forecast_ETS_ConfInt** supports up to 30% missing data, and automatically adjusts for it. See Remarks.|
| _Arg7_|Optional|VARIANT|Aggregation: Although the timeline requires a constant step between data points,  **Forecast_ETS_ConfInt** aggregates multiple points that have the same time stamp. See Remarks.|

### Return Value

 **Double**


## Remarks

 It isn't necessary to sort the timeline ( _Arg3_), because  **Forecast_ETS_ConfInt** sorts it implicitly for calculations. If **Forecast_ETS_ConfInt** can't identify a constant step in the timeline, it returns runtime error ?1004?. If the timeline contains duplicate values, **Forecast_ETS_ConfInt** also returns an error. If the ranges of the timeline and values aren't all of the same size, **Forecast_ETS_ConfInt** returns runtime error ?1004?.

A confidence interval ( _Arg4_) of 95% means that 95% of future points are expected to fall within this radius from the result [Forecast_ETS](worksheetfunction-forecast_ets-method-excel.md) forecasted (with normal distribution). Using confidence intervals can help you grasp the accuracy of the predicted model. A smaller interval implies more confidence in the prediction for this specific point.

For example, for a 90% confidence interval, a 90% confidence level is computed (90% of future points are to fall within this radius from prediction). The default value is 95%. For numbers outside of the range (0,1),  **Forecast_ETS_ConfInt** returns an error.

The default value of 1 for seasonality ( _Arg5_) means Excel detects seasonality automatically for the forecast and uses positive, whole numbers for the length of the seasonal pattern. 0 indicates no seasonality, meaning the prediction will be linear. Positive whole numbers indicate to the algorithm to use patterns of this length as the seasonality. For any other value,  **Forecast_ETS_ConfInt** returns an error. Maximum supported seasonality is 8,760 (the number of hours in a year). Any seasonality value above that number results in an error.

Passing 0 for the data completions parameter ( _Arg6_) instructs the algorithm to account for missing points as zeros. The default value of 1 accounts for missing points by computing them to be the average of the neighboring points. If there is more than 30% missing data,  **Forecast_ETS_ConfInt** returns runtime error ?1004?.

The aggregation parameter ( _Arg7_) is a numeric value specifying the method to use to aggregate several values that have the same time stamp. The default value of 0 specifies AVERAGE, while other numbers between 1 and 6 specify SUM, COUNT, COUNTA, MIN, MAX, and MEDIAN.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

