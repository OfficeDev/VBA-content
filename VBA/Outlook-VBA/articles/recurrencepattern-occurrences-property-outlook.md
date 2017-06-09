---
title: RecurrencePattern.Occurrences Property (Outlook)
keywords: vbaol11.chm282
f1_keywords:
- vbaol11.chm282
ms.prod: outlook
api_name:
- Outlook.RecurrencePattern.Occurrences
ms.assetid: a99a8a1c-dcd3-e96d-6091-0a005ca3b05f
ms.date: 06/08/2017
---


# RecurrencePattern.Occurrences Property (Outlook)

Returns or sets a  **Long** indicating the number of occurrences of the recurrence pattern. Read/write.


## Syntax

 _expression_ . **Occurrences**

 _expression_ A variable that represents a **RecurrencePattern** object.


## Remarks

This property allows the definition of a recurrence pattern that is only valid for the specified number of subsequent occurrences. For example, you can set this property to 10 for a formal training course that will be held on the next ten Thursday evenings. This property must be coordinated with other properties when setting up a recurrence pattern. If the  **[PatternEndDate](recurrencepattern-patternenddate-property-outlook.md)** property or the **Occurrences** property is set, the pattern is considered to be finite and the **[NoEndDate](recurrencepattern-noenddate-property-outlook.md)** property is **False** . If neither **PatternEndDate** nor **Occurrences** is set, the pattern is considered infinite and **NoEndDate** is **True** .


## See also


#### Concepts


[RecurrencePattern Object](recurrencepattern-object-outlook.md)

