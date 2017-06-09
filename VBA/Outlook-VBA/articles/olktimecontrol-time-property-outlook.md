---
title: OlkTimeControl.Time Property (Outlook)
keywords: vbaol11.chm1000393
f1_keywords:
- vbaol11.chm1000393
ms.prod: outlook
api_name:
- Outlook.OlkTimeControl.Time
ms.assetid: da483b8b-ef16-53e6-b3a8-e18f71799759
ms.date: 06/08/2017
---


# OlkTimeControl.Time Property (Outlook)

Returns or sets a  **Date** that represents the time value currently selected in the control. Read/write.


## Syntax

 _expression_ . **Time**

 _expression_ A variable that represents an **OlkTimeControl** object.


## Remarks

The default value is Dec 30, 1899 12:00 AM.

When using the time control to indicate a duration (that is, the  **[Style](olktimecontrol-style-property-outlook.md)** is **olTimeStyleDuration** ), if the duration is longer than 24 hours, the **Time** property will indicate the duration from Dec 30, 1899 12:00 AM. For example, a duration spanning 24 hours will return a **Date** value of Dec 31 1899 12:00 AM. If this is a duration value for an appointment and you would like to determine an end time for the appointment, you can add this value to the **[ReferenceTime](olktimecontrol-referencetime-property-outlook.md)** property value.


## See also


#### Concepts


[OlkTimeControl Object](olktimecontrol-object-outlook.md)

