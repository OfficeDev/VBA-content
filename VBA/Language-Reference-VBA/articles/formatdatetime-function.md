---
title: FormatDateTime Function
keywords: vblr6.chm1011367
f1_keywords:
- vblr6.chm1011367
ms.prod: office
ms.assetid: 1ead64ea-cea4-0464-a6e4-f28b1edb06cc
ms.date: 06/08/2017
---


# FormatDateTime Function



 **Description**
Returns an expression formatted as a date or time.
 **Syntax**
 **FormatDateTime(**_Date_ [ **,**_NamedFormat_ ] **)**
The  **FormatDateTime** function syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _Date_|Required. Date expression to be formatted.|
| _NamedFormat_|Optional. Numeric value that indicates the date/time format used. If omitted,  **vbGeneralDate** is used.|
 **Settings**
The  _NamedFormat_ argument has the following settings:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbGeneralDate**|0|Display a date and/or time. If there is a date part, display it as a short date. If there is a time part, display it as a long time. If present, both parts are displayed.|
|**vbLongDate**|1|Display a date using the long date format specified in your computer's regional settings.|
|**vbShortDate**|2|Display a date using the short date format specified in your computer's regional settings.|
|**vbLongTime**|3|Display a time using the time format specified in your computer's regional settings.|
|**vbShortTime**|4|Display a time using the 24-hour format (hh:mm).|

