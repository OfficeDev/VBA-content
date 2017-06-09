---
title: Document.Time Property (Visio)
keywords: vis_sdr.chm10550900
f1_keywords:
- vis_sdr.chm10550900
ms.prod: visio
api_name:
- Visio.Document.Time
ms.assetid: 04d7d5d9-9e4f-c64a-faa9-ac521807b44f
ms.date: 06/08/2017
---


# Document.Time Property (Visio)

Returns the most recently recorded date and time. Read-only.


## Syntax

 _expression_ . **Time**

 _expression_ A variable that represents a **Document** object.


### Return Value

Date


## Remarks

The  **Time** property is updated whenever any values are updated in the following: the **TimeEdited** property, the **TimePrinted** property, the **TimeCreated** property, the **TimeSaved** property, or the NOW function.

In the  **Date** type, the value to the left of the decimal point represents the date, and the value to the right of the decimal point represents the time. For example, the **Date** value 38000.75 represents 6:00 P.M. on January 14, 2004.

If you convert a  **Date** value to the **String** type, the date is rendered according to the short date format recognized by your computer. Times are displayed according to the time format (either 12-hour or 24-hour) recognized by your computer.


