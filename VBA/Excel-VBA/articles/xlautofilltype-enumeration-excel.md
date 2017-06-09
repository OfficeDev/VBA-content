---
title: XlAutoFillType Enumeration (Excel)
ms.prod: excel
api_name:
- Excel.XlAutoFillType
ms.assetid: bfb09be7-8922-ef4b-751e-c8972536b723
ms.date: 06/08/2017
---


# XlAutoFillType Enumeration (Excel)

Specifies how the target range is to be filled, based on the contents of the source range.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **xlFillCopy**|1|Copy the values and formats from the source range to the target range, repeating if necessary.|
| **xlFillDays**|5|Extend the names of the days of the week in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary.|
| **xlFillDefault**|0|Excel determines the values and formats used to fill the target range.|
| **xlFillFormats**|3|Copy only the formats from the source range to the target range, repeating if necessary.|
| **xlFillMonths**|7|Extend the names of the months in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary.|
| **xlFillSeries**|2|Extend the values in the source range into the target range as a series (for example, '1, 2' is extended as '3, 4, 5'). Formats are copied from the source range to the target range, repeating if necessary.|
| **xlFillValues**|4|Copy only the values from the source range to the target range, repeating if necessary.|
| **xlFillWeekdays**|6|Extend the names of the days of the workweek in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary.|
| **xlFillYears**|8|Extend the years in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary.|
| **xlGrowthTrend**|10|Extend the numeric values from the source range into the target range, assuming that the relationships between the numbers in the source range are multiplicative (for example, '1, 2,' is extended as '4, 8, 16', assuming that each number is a result of multiplying the previous number by some value). Formats are copied from the source range to the target range, repeating if necessary.|
| **xlLinearTrend**|9|Extend the numeric values from the source range into the target range, assuming that the relationships between the numbers is additive (for example, '1, 2,' is extended as '3, 4, 5', assuming that each number is a result of adding some value to the previous number). Formats are copied from the source range to the target range, repeating if necessary.|
| **xlFlashFill**|11|Extend the values from the source range into the target range based on the detected pattern of previous user actions, repeating if necessary.|

