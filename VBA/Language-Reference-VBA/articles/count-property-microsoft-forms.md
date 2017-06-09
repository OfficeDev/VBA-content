---
title: Count Property (Microsoft Forms)
keywords: fm20.chm2001000
f1_keywords:
- fm20.chm2001000
ms.prod: office
ms.assetid: 84580b94-05da-57d9-780b-e95545a5ea37
ms.date: 06/08/2017
---


# Count Property (Microsoft Forms)



Returns the number of objects in a [collection](vbe-glossary.md).
 **Syntax**
 _object_. **Count**
The  **Count** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Remarks**
The  **Count** property is read only.
Note that the index value for the first page or tab of a collection is zero, the value for the second page or tab is one, and so on. For example, if a  **MultiPage** contains two pages, the indexes of the pages are 0 and 1, and the value of **Count** is 2.

