---
title: RecurrencePattern.PatternEndDate Property (Outlook)
keywords: vbaol11.chm283
f1_keywords:
- vbaol11.chm283
ms.prod: outlook
api_name:
- Outlook.RecurrencePattern.PatternEndDate
ms.assetid: 0f78ea71-3d92-2d38-be10-e05ab7bcf44a
ms.date: 06/08/2017
---


# RecurrencePattern.PatternEndDate Property (Outlook)

Returns or sets a  **Date** indicating the end date for the recurrence pattern. Read/write.


## Syntax

 _expression_ . **PatternEndDate**

 _expression_ A variable that represents a **RecurrencePattern** object.


## Remarks

This property is optional but must be coordinated with other properties when setting up a recurrence pattern. If this property or the  **[Occurrences](recurrencepattern-occurrences-property-outlook.md)** property is set, the pattern is considered to be finite, and the **[NoEndDate](recurrencepattern-noenddate-property-outlook.md)** property is **False** . If neither **PatternEndDate** nor **Occurrences** is set, the pattern is considered infinite and **NoEndDate** is **True** . The **[Interval](recurrencepattern-interval-property-outlook.md)** property must be set before setting **PatternEndDate** .


## See also


#### Concepts


[RecurrencePattern Object](recurrencepattern-object-outlook.md)

