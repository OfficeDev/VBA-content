---
title: PivotField.ClearLabelFilters Method (Excel)
keywords: vbaxl10.chm240156
f1_keywords:
- vbaxl10.chm240156
ms.prod: excel
api_name:
- Excel.PivotField.ClearLabelFilters
ms.assetid: 48b8f6be-b4c0-26c6-2550-63729fd6a918
ms.date: 06/08/2017
---


# PivotField.ClearLabelFilters Method (Excel)

This method deletes all label filters or all date filters in the  **PivotFilters** collection of the PivotField.


## Syntax

 _expression_ . **ClearLabelFilters**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

The following table lists the different label filter types that will be deleted by this method.



||
| **xlCaptionEquals**|
| **xlCaptionDoesNotEqual**|
| **xlCaptionIsGreaterThan**|
| **xlCaptionIsGreaterThanOrEqualTo**|
| **xlCaptionIsLessThan**|
| **xlCaptionIsLessThanOrEqualTo**|
| **xlCaptionBeginsWith**|
| **xlCaptionDoesNotBeginWith**|
| **xlCaptionEndsWith**|
| **xlCaptionDoesNotEndWith**|
| **xlCaptionContains**|
| **xlCaptionDoesNotContain**|
| **xlCaptionIsBetween**|
| **xlCaptionIsNotBetween**|
||
||
The following table lists the different date filter types that will be deleted by this method.



||
| <strong>xlSpecificDate</strong>|
| 
<strong>xlNotSpecificDate</strong>|
| 
<strong>xlBefore</strong>|
| 
<strong>xlBeforeOrEqualTo</strong>|
| 
<strong>xlAfter</strong>|
| 
<strong>xlAfterOrEqualTo</strong>|
| 
<strong>xlDateBetween</strong>|
| 
<strong>xlDateNotBetween</strong>|
| 
<strong>xlDateToday</strong>|
| 
<strong>xlDateYesterday</strong>|
| 
<strong>xlDateTomorrow</strong>|
| 
<strong>xlDateThisWeek</strong>|
| 
<strong>xlDateLastWeek</strong>|
| 
<strong>xlDateNextWeek</strong>|
| 
<strong>xlDateThisMonth</strong>|
| 
<strong>xlDateLastMonth</strong>|
| 
<strong>xlDateNextMonth</strong>|
| 
<strong>xlDateThisQuarter</strong>|
| 
<strong>xlDateLastQuarter</strong>|
| 
<strong>xlDateNextQuarter</strong>|
| 
<strong>xlDateThisYear</strong>|
| 
<strong>xlDateLastYear</strong>|
| 
<strong>xlDateNextYear</strong>|
| 
<strong>xlYearToDate</strong>|
| 
<strong>xlAllDatesInPeriodQuarter1</strong>|
| 
<strong>xlAllDatesInPeriodQuarter2</strong>|
| 
<strong>xlAllDatesInPeriodQuarter3</strong>|
| 
<strong>xlAllDatesInPeriodQuarter4</strong>|
| 
<strong>xlAllDatesInPeriodJanuary</strong>|
| 
<strong>xlAllDatesInPeriodFebruary</strong>|
| 
<strong>xlAllDatesInPeriodMarch</strong>|
| 
<strong>xlAllDatesInPeriodApril</strong>|
| 
<strong>xlAllDatesInPeriodMay</strong>|
| 
<strong>xlAllDatesInPeriodJune</strong>|
| 
<strong>xlAllDatesInPeriodJuly</strong>|
| 
<strong>xlAllDatesInPeriodAugust</strong>|
| 
<strong>xlAllDatesInPeriodSeptember</strong>|
| 
<strong>xlAllDatesInPeriodOctober</strong>|
| 
<strong>xlAllDatesInPeriodNovember</strong>|
| 
<strong>xlAllDatesInPeriodDecember</strong>|

## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

