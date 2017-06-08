---
title: Remove Method
keywords: fm20.chm2000360
f1_keywords:
- fm20.chm2000360
ms.prod: office
api_name:
- Office.Remove
ms.assetid: 16ee4145-3e1e-9e44-7af1-2ecd3a92c9e3
ms.date: 06/08/2017
---


# Remove Method



Removes a member from a [collection](vbe-glossary.md); or, removes a control from a  **Frame**, **Page**, or form.
 **Syntax**
 _object_. **Remove(**_collectionindex_**)**
The  **Remove** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _collectionindex_|Required. A member's position, or index, within a collection. Numeric as well as string values are acceptable. If the value is a number, the minimum value is zero, and the maximum value is one less than the number of members in the collection. If the value is a string, it must correspond to a valid member name.|
 **Remarks**
This method deletes any control that was added at [run time](vbe-glossary.md). However, attempting to delete a control that was added at [design time](vbe-glossary.md) will result in an error.

