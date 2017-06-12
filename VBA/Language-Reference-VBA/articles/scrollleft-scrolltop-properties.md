---
title: ScrollLeft, ScrollTop Properties
keywords: fm20.chm5225086
f1_keywords:
- fm20.chm5225086
ms.prod: office
ms.assetid: 1b60c64d-84e5-6e21-eebf-a4c375e7c148
ms.date: 06/08/2017
---


# ScrollLeft, ScrollTop Properties



Specify the distance, in [points](vbe-glossary.md), of the left or top edge of the visible form from the left or top edge of the logical form, page, or control.
 **Syntax**
 _object_. **ScrollLeft** [= _Single_ ]
 _object_. **ScrollTop** [= _Single_ ]
The  **ScrollLeft** and **ScrollTop** property syntaxes have these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Single_|Optional. The distance from the edge of the form.|
 **Remarks**
The minimum value is zero; the maximum value is the difference between the value of the  **ScrollWidth** property and the value of the **Width** property for the form or page.

