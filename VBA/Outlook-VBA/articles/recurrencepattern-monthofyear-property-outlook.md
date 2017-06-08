---
title: RecurrencePattern.MonthOfYear Property (Outlook)
keywords: vbaol11.chm280
f1_keywords:
- vbaol11.chm280
ms.prod: outlook
api_name:
- Outlook.RecurrencePattern.MonthOfYear
ms.assetid: 14112950-1e2a-a99a-7c48-3e76358de645
ms.date: 06/08/2017
---


# RecurrencePattern.MonthOfYear Property (Outlook)

Returns or sets a  **Long** indicating which month of the year is valid for the specified recurrence pattern. Read/write.


## Syntax

 _expression_ . **MonthOfYear**

 _expression_ A variable that represents a **RecurrencePattern** object.


## Remarks

The value can be a number from 1 through 12. For example, setting this property to 5 and the  **[RecurrenceType](recurrencepattern-recurrencetype-property-outlook.md)** property to **olRecursYearly** would cause this recurrence pattern to occur every May. This property is only valid for recurrence patterns whose **RecurrenceType** property is set to **olRecursYearly** or **olRecursYearNth** .


## See also


#### Concepts


[RecurrencePattern Object](recurrencepattern-object-outlook.md)

