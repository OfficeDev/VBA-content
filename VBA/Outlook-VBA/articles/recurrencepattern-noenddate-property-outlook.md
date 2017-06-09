---
title: RecurrencePattern.NoEndDate Property (Outlook)
keywords: vbaol11.chm281
f1_keywords:
- vbaol11.chm281
ms.prod: outlook
api_name:
- Outlook.RecurrencePattern.NoEndDate
ms.assetid: 47c5841a-c0d2-2b06-ec73-7093779ceafa
ms.date: 06/08/2017
---


# RecurrencePattern.NoEndDate Property (Outlook)

Returns a  **Boolean** value that indicates whether the recurrence pattern has no end date. Read/write.


## Syntax

 _expression_ . **NoEndDate**

 _expression_ A variable that represents a **RecurrencePattern** object.


## Remarks

This property must be coordinated with other properties when setting up a recurrence pattern. If the  **[PatternEndDate](recurrencepattern-patternenddate-property-outlook.md)** property or the **[Occurrences](recurrencepattern-occurrences-property-outlook.md)** property is set, the pattern is considered to be finite and the **NoEndDate** property is **False** . If neither **PatternEndDate** nor **Occurrences** is set, the pattern is considered infinite and **NoEndDate** is **True** .


## See also


#### Concepts


[RecurrencePattern Object](recurrencepattern-object-outlook.md)

