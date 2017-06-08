---
title: OlkTimeControl.IntervalTime Property (Outlook)
keywords: vbaol11.chm1000397
f1_keywords:
- vbaol11.chm1000397
ms.prod: outlook
api_name:
- Outlook.OlkTimeControl.IntervalTime
ms.assetid: 518bd878-f970-2600-8c63-17fa8774def4
ms.date: 06/08/2017
---


# OlkTimeControl.IntervalTime Property (Outlook)

Returns or sets a  **Date** that specifies the number of minutes displayed as an interval used for the **olTimeStyleTimeDuration** style setting on the time control. Read/write.


## Syntax

 _expression_ . **IntervalTime**

 _expression_ A variable that represents an **OlkTimeControl** object.


## Remarks

The default value is 30.

The minimum value for  **IntervalTime** is 1 and the maximum value is 1440. Assigning a value outside of this range will result in the nearest edge value being used instead.

When  **[Style](olktimecontrol-style-property-outlook.md)** is **olTimeStyleTimeDuration** , the date control displays the value of **[ReferenceTime](olktimecontrol-referencetime-property-outlook.md)** as the first selectable time value, shows additional intervals (specified by **IntervalTime** ) starting from the **ReferenceTime** value, and shows the duration of an event.

The default value for  **ReferenceTime** is 30 Dec 1899 12:00 AM. In this case, the time control will display **30 Dec 1899 12:00 AM** as the first selectable time. A value of 60 for **IntervalTime** will mark **30 Dec 1899 1:00 AM** as the first interval.


## See also


#### Concepts


[OlkTimeControl Object](olktimecontrol-object-outlook.md)

