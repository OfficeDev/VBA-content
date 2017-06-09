---
title: PivotFilter.FilterType Property (Excel)
keywords: vbaxl10.chm770074
f1_keywords:
- vbaxl10.chm770074
ms.prod: excel
api_name:
- Excel.PivotFilter.FilterType
ms.assetid: 0c7b7a0c-1472-8a36-2876-62353568ec90
ms.date: 06/08/2017
---


# PivotFilter.FilterType Property (Excel)

Specifies the type of filter to be applied. Read-only  **xlPivotFilterType** .


## Syntax

 _expression_ . **FilterType**

 _expression_ A variable that represents a **PivotFilter** object.


## Remarks

The following table shows the new filter types, the arguments that must be supplied, and the arguments that cannot be supplied (Unavailable) for each filter type.



|**xlPivotFilterType**|**DataField**|**DataCubeField (OLAP)**|**Value1**|**Value2**|
|:-----|:-----|:-----|:-----|:-----|
|xlTopCount|Required|Required|Required|Unavailable|
|xlBottomCount|Required|Required|Required|Unavailable|
|xlTopPercent|Required|Required|Required|Unavailable|
|xlBottomPercent|Required|Required|Required|Unavailable|
|xlTopSum|Required|Required|Required|Unavailable|
|xlBottomSum|Required|Required|Required|Unavailable|
|xlCaptionEquals|Unavailable|Unavailable|Required|Unavailable|
|xlCaptionDoesNotEqual|Unavailable|Unavailable|Required|Unavailable|
|xlCaptionIsGreaterThan|Unavailable|Unavailable|Required|Unavailable|
|xlCaptionIsGreaterThanOrEqualTo|Unavailable|Unavailable|Required|Unavailable|
|xlCaptionIsLessThan|Unavailable|Unavailable|Required|Unavailable|
|xlCaptionIsLessThanOrEqualTo|Unavailable|Unavailable|Required|Unavailable|
|xlCaptionBeginsWith|Unavailable|Unavailable|Required|Unavailable|
|xlCaptionDoesNotBeginWith|Unavailable|Unavailable|Required|Unavailable|
|xlCaptionEndsWith|Unavailable|Unavailable|Required|Unavailable|
|xlCaptionDoesNotEndWith|Unavailable|Unavailable|Required|Unavailable|
|xlCaptionContains|Unavailable|Unavailable|Required|Unavailable|
|xlCaptionDoesNotContain|Unavailable|Unavailable|Required|Unavailable|
|xlCaptionIsBetween|Unavailable|Unavailable|Required|Required|
|xlCaptionIsNotBetween|Unavailable|Unavailable|Required|Required|
|xlValueEquals|Required|Required|Required|Unavailable|
|xlValueDoesNotEqual|Required|Required|Required|Unavailable|
|xlValueIsGreaterThan|Required|Required|Required|Unavailable|
|xlValueIsGreaterThanOrEqualTo|Required|Required|Required|Unavailable|
|xlValueIsLessThan|Required|Required|Required|Unavailable|
|xlValueIsLessThanOrEqualTo|Required|Required|Required|Unavailable|
|xlValueIsBetween|Required|Required|Required|Required|
|xlValueIsNotBetween|Required|Required|Required|Required|
|xlSpecificDate|Unavailable|Unavailable|Required|Unavailable|
|xlNotSpecificDate|Unavailable|Unavailable|Required|Unavailable|
|xlBefore|Unavailable|Unavailable|Required|Unavailable|
|xlBeforeOrEqualTo|Unavailable|Unavailable|Required|Unavailable|
|xlAfter|Unavailable|Unavailable|Required|Unavailable|
|xlAfterOrEqualTo|Unavailable|Unavailable|Required|Unavailable|
|xlBetween|Unavailable|Unavailable|Required|Unavailable|
|xlNotBetween|Unavailable|Unavailable|Required|Unavailable|
|xlFilterToday|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterYesterday|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterTomorrow|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterThisWeek|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterLastWeek|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterNextWeek|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterThisMonth|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterLastMonth|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterNextMonth|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterThisQuarter|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterLastQuarter|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterNextQuarter|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterThisYear|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterLastYear|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterNextYear|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterYearToDate|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterAllDatesInPeriodQuarter1|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterAllDatesInPeriodQuarter2|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterAllDatesInPeriodQuarter3|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterAllDatesInPeriodQuarter4|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterAllDatesInPeriodJanuary|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterAllDatesInPeriodFebruary|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterAllDatesInPeriodMarch|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterAllDatesInPeriodApril|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterAllDatesInPeriodMay|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterAllDatesInPeriodJune|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterAllDatesInPeriodJuly|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterAllDatesInPeriodAugust|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterAllDatesInPeriodSeptember|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterAllDatesInPeriodOctober|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterAllDatesInPeriodNovember|Unavailable|Unavailable|Unavailable|Unavailable|
|xlFilterAllDatesInPeriodDecember|Unavailable|Unavailable|Unavailable|Unavailable|

## See also


#### Concepts


[PivotFilter Object](pivotfilter-object-excel.md)

