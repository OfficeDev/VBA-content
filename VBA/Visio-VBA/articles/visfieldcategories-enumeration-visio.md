---
title: VisFieldCategories Enumeration (Visio)
keywords: vis_sdr.chm70135
f1_keywords:
- vis_sdr.chm70135
ms.prod: visio
ms.assetid: f10df918-5be3-e883-1da5-2a932fd1074f
ms.date: 06/08/2017
---


# VisFieldCategories Enumeration (Visio)

Text-field categories passed to the  **AddField** and **AddFieldEx** methods.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visFCatCustom**|0|Custom field.|
| **visFCatDateTime**|1|Date/time field.|
| **visFCatDocument**|2|Document information field.|
| **visFCatGeometry**|3|Geometery field.|
| **visFCatObject**|4|Object information.|
| **visFCatPage**|5|Page information field.|

 **Note**  Do not pass  **visFCatCustom** to the **AddField** method. Pass it to **AddCustomField** or **AddCustomFieldU** instead.


