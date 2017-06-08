---
title: WorksheetFunction.Forecast_ETS_STAT Method (Excel)
keywords: vbaxl10.chm137472
f1_keywords:
- vbaxl10.chm137472
ms.assetid: 6b1c0256-3146-4dc5-3f8a-27e61a982fee
ms.date: 06/08/2017
ms.prod: excel
---


# WorksheetFunction.Forecast_ETS_STAT Method (Excel)

Returns a statistical value as a result of time series forecasting.


## Syntax

 _expression_ . **Forecast_ETS_STAT**( _Arg1_,  _Arg2_,  _Arg3_,  _Arg4_,  _Arg5_,  _Arg6_)

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|||||
| _Arg1_|Required|VARIANT|Values: the historical values, for which you want to forecast the next points.|
| _Arg2_|Required|VARIANT|Timeline: the independent array or range of dates or numeric data. The values in the timeline must have a consistent step between them and can?t be zero. See Remarks.|
| _Arg3_|Required|DOUBLE|Statistic_type:??? A numeric value between 1 and 8, indicating which statistic will be returned for the calculated forecast.|
| _Arg4_|Optional|VARIANT|Confidence level: A numerical value between 0 and 1 (exclusive), indicating a confidence level for the calculated confidence interval. See Remarks.|
| _Arg5_|Optional|VARIANT|Data completions: Although the timeline requires a constant step between data points,  **Forecast_ETS_STAT** supports up to 30% missing data, and automatically adjusts for it. See Remarks.|
| _Arg6_|Optional|VARIANT|Aggregation: Although the timeline requires a constant step between data points,  **Forecast_ETS_STAT** aggregates multiple points that have the same time stamp. See Remarks.|

### Return Value

 **DOUBLE**


## Remarks

 It isn't necessary to sort the timeline ( _Arg2_), because  **Forecast_ETS_STAT** sorts it implicitly for calculations. If **Forecast_ETS_STAT** can't identify a constant step in the timeline, it returns runtime error ?1004?. If the timeline contains duplicate values, **Forecast_ETS_STAT** also returns an error. If the ranges of the timeline and values aren't all of the same size, **Forecast_ETS_STAT** returns runtime error ?1004?.

The statistic_type parameter ( _Arg3_) indicates which statistic is requested by this function. The following optional statistics can be returned:


-  **Alpha parameter** of ETS algorithm. Returns the base value parameter?a higher value gives more weight to recent data points.
    
-  **Beta** parameter of ETS algorithm. Returns the trend value parameter?a higher value gives more weight to the recent trend.
    
-  **Gamma** parameter of ETS algorithm. Returns the trend value parameter?a higher value gives more weight to the recent trend.
    
-  **MASE** metric. Returns the mean absolute scaled error metric?a measure of the accuracy of forecasts.
    
-  **SMAPE** metric. Returns the symmetric mean absolute percentage error metric?an accuracy measure based on percentage errors.
    
-  **MAE** metric. Returns the symmetric mean absolute percentage error metric?an accuracy measure based on percentage errors.
    
-  **RMSE** metric. Returns the root mean squared error metric?a measure of the differences between predicted and observed values.
    
-  **Step size detected** .? Returns the step size detected in the historical timeline.
    
A confidence interval ( _Arg4_) of 95% means that 95% of future points are expected to fall within this radius from the result [Forecast_ETS](worksheetfunction-forecast_ets-method-excel.md) forecasted (with normal distribution). Using confidence intervals can help you grasp the accuracy of the predicted model. A smaller interval implies more confidence in the prediction for this specific point.

For example, for a 90% confidence interval, a 90% confidence level is computed (90% of future points are to fall within this radius from prediction). The default value is 95%. For numbers outside of the range (0,1),  **Forecast_ETS_STAT** returns an error.

Passing 0 for the data completions parameter ( _Arg5_) instructs the algorithm to account for missing points as zeros. The default value of 1 accounts for missing points by computing them to be the average of the neighboring points. If there is more than 30% missing data,  **Forecast_ETS_STAT** returns runtime error ?1004?.

The aggregation parameter ( _Arg6_) is a numeric value specifying the method to use to aggregate several values that have the same time stamp. The default value of 0 specifies AVERAGE, while other numbers between 1 and 6 specify SUM, COUNT, COUNTA, MIN, MAX, and MEDIAN.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

