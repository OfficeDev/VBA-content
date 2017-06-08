---
title: Document.TimeSaved Property (Visio)
keywords: vis_sdr.chm10550920
f1_keywords:
- vis_sdr.chm10550920
ms.prod: visio
api_name:
- Visio.Document.TimeSaved
ms.assetid: 801c7940-b838-15ae-cee8-e07ca5ae78ea
ms.date: 06/08/2017
---


# Document.TimeSaved Property (Visio)

Returns the date and time the document was last saved. Read-only.


## Syntax

 _expression_ . **TimeSaved**

 _expression_ A variable that represents a **Document** object.


### Return Value

Date


## Remarks

In the  **Date** type, the value to the left of the decimal point represents the date, and the value to the right of the decimal point represents the time. For example, the **Date** value 38000.75 represents 6:00 P.M. on January 14, 2004.

If you convert a  **Date** value to the **String** type, the date is rendered according to the short date format recognized by your computer. Times are displayed according to the time format (either 12-hour or 24-hour) recognized by your computer.


