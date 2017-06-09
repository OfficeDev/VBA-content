---
title: Calendar Property
keywords: vblr6.chm1117202
f1_keywords:
- vblr6.chm1117202
ms.prod: office
api_name:
- Office.Calendar
ms.assetid: ca321712-934e-2aee-46b8-b2895be362ea
ms.date: 06/08/2017
---


# Calendar Property



Returns or sets a value specifying the type of calendar to use with your [project](vbe-glossary.md).
You can use one of two settings for  **Calendar**:


|**Setting**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbCalGreg**|0|Use Gregorian calendar (default).|
|**vbCalHijri**|1|Use Hijri calendar.|
 **Remarks**
You can only set the  **Calendar** property programmatically. For example, to use the Hijri calendar, use:



```
Calendar = vbCalHijri


```


