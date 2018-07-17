---
title: Document.TimeCreated Property (Visio)
keywords: vis_sdr.chm10550905
f1_keywords:
- vis_sdr.chm10550905
ms.prod: visio
api_name:
- Visio.Document.TimeCreated
ms.assetid: efb0fdc6-c4ff-78a5-08bb-7a4367cedc43
ms.date: 06/08/2017
---


# Document.TimeCreated Property (Visio)

Returns the date and time the document was created. Read-only.


## Syntax

 _expression_ . **TimeCreated**

 _expression_ A variable that represents a **Document** object.


### Return Value

Date


## Remarks

In the  **Date** type, the value to the left of the decimal point represents the date, and the value to the right of the decimal point represents the time. For example, the **Date** value 38000.75 represents 6:00 P.M. on January 14, 2004.

If you convert a  **Date** value to the **String** type, the date is rendered according to the short date format recognized by your computer. Times are displayed according to the time format (either 12-hour or 24-hour) recognized by your computer.


