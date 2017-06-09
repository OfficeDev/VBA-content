---
title: Workbook.CreateForecastSheet Method (Excel)
keywords: vbaxl10.chm199280
f1_keywords:
- vbaxl10.chm199280
ms.assetid: bec7b60b-7840-af15-6d5f-f5c184ea7aee
ms.date: 06/08/2017
ms.prod: excel
---


# Workbook.CreateForecastSheet Method (Excel)

If you have historical time-based data, you can use  **CreateForecastSheet** to create a forecast. When you create a forecast, a new worksheet is created that contains a table of the historical and predicted values and a chart showing this. A forecast can help you predict things like future sales, inventory requirements, or consumer trends.


## Syntax

 _expression_ . **CreateForecastSheet**( _Timeline_,  _Timeline_,  _Values_,  _ForecastStart_,  _ForecastEnd_,  _ConfInt_,  _Seasonality_,  _DataCompletion_,  _Aggregation_,  _ChartType_,  _ShowStatsTable_)

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|||||
| _Timeline_|Required|RANGE|The independent array or range of numeric data. The dates in the timeline must have a consistent step between them and can?t be zero. The timeline isn't required to be sorted, as the forecast mechanism will sort it implicitly for calculations. If a constant step can't be identified in the provided timeline, then Invalid procedure call or argument (Error 5) will be returned.|
| _Values_|Required|RANGE| Values are the historical values, for which you want to forecast the next points.|
| _ForecastStart_|Optional|VARIANT|The point from which the generated forecast will begin.|
| _ForecastEnd_|Optional|VARIANT|The point in which the generated forecast will end.|
| _ConfInt_|Optional|VARIANT|A numerical value between 0 and 1 (exclusive), indicating a confidence level for the calculated confidence interval. For example, for a 90% confidence interval, a 90% confidence level will be computed (90% of future points are to fall within this radius from prediction). The default value is 95%.|
| _Seasonality_|Optional|VARIANT|A numeric value. The default value of 1 means Excel detects seasonality automatically for the forecast and uses positive, whole numbers for the length of the seasonal pattern. 0 indicates no seasonality, meaning the prediction will be linear. Positive whole numbers will indicate to the algorithm to use patterns of this length as the seasonality. For any other value, Error 5 will be returned. Maximum supported seasonality is 8,760 (number of hours in a year). Any seasonality above that number will result in the Error 5.|
| _DataCompletion_|Optional|VARIANT| _DataCompletion_ can be one of these:[XlForecastDataCompletion](xlforecastdatacompletion-enumeration-excel.md) constants: **xlDataCompletionZeros** or **xlDataCompletionInterpolate** . The default is **xlDataCompletionInterpolate** .|
| _Aggregation_|Optional|VARIANT| _Aggregation_ can be one of these[XlForecastAggregation](xlforecastaggregation-enumeration-excel.md) constants: **xlAggregationAverage** , **xlAggregationCount** , **xlAggregationCountA** , **xlAggregationMax** , **xlAggregationMedian** , **xlAggregationMin** or **xlAggregationSum** . The default is **xlAggregationAverage** .|
| _ChartType_|Optional|VARIANT| _ChartType_ can be one of these[XlForecastChartType](xlforecastcharttype-enumeration-excel.md) constants: **xlChartTypeLine** or **xlChartTypeColumn** . The default is **xlChartTypeLine** .|
| _ShowStatsTable_|Optional|VARIANT| **True** or **False** . If **True**, an additional table is generated in the created sheet. This table contains statistical measures that indicate the accuracy of the created forecast.|

### Return Value

None


## Remarks

When you use a formula to create a forecast, it returns a table with the historical and predicted data, and a chart. The forecast predicts future values using your existing time-based data and the AAA version of the Exponential Smoothing (ETS) algorithm. The table has the following columns, three of which are calculated columns:


- Historical time column (your time-based data series)
    
- Historical values column (your corresponding values data series)
    
- Forecasted values column (calculated using FORECAST_ETS)
    
- Two columns representing the confidence interval (calculated using FORECAST_ETS_CONFINT)
    

## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

