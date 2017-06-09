---
title: RecurrencePattern.Instance Property (Outlook)
keywords: vbaol11.chm278
f1_keywords:
- vbaol11.chm278
ms.prod: outlook
api_name:
- Outlook.RecurrencePattern.Instance
ms.assetid: 3458aeff-97b7-02f8-e352-203ecc92dedd
ms.date: 06/08/2017
---


# RecurrencePattern.Instance Property (Outlook)

Returns or sets a  **Long** specifying the count for which the recurrence pattern is valid for a given interval. Read/write.


## Syntax

 _expression_ . **Instance**

 _expression_ A variable that represents a **RecurrencePattern** object.


## Remarks

This property is only valid for recurrences of the  **olRecursMonthNth** and **olRecursYearNth** type and allows the definition of a recurrence pattern that is only valid for the Nth occurrence, such as "the 2nd Sunday in March" pattern. The count is set numerically: 1 for the first, 2 for the second, and so on through 5 for the last. Values greater than 5 will generate errors when the pattern is saved.


## See also


#### Concepts


[RecurrencePattern Object](recurrencepattern-object-outlook.md)

