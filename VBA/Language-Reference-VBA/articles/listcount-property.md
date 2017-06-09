---
title: ListCount Property
keywords: fm20.chm5225054
f1_keywords:
- fm20.chm5225054
ms.prod: office
api_name:
- Office.ListCount
ms.assetid: e6878930-514c-47cb-0961-bd9f5f79caff
ms.date: 06/08/2017
---


# ListCount Property



Returns the number of list entries in a control.
 **Syntax**
 _object_. **ListCount**
The  **ListCount** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
|||
| _object_|Required. A valid object.|
 **Remarks**
The  **ListCount** property is read-only. **ListCount** is the number of rows over which you can scroll. **ListRows** is the maximum to display at once. **ListCount** is always one greater than the largest value for the **ListIndex** property, because index numbers begin with 0 and the count of items begins with 1. If no item is selected, **ListCount** is 0 and **ListIndex** is -1.

