---
title: Item Method (Microsoft Forms)
keywords: fm20.chm5224962
f1_keywords:
- fm20.chm5224962
ms.prod: office
ms.assetid: 6b50b145-7598-157d-111c-5ba9234520bd
ms.date: 06/08/2017
---


# Item Method (Microsoft Forms)



Returns a member of a [collection](vbe-glossary.md), either by position or by name.
 **Syntax**
 **Set**_Object_ = _object_. **Item(**_collectionindex_**)**
The  **Item** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _collectionindex_|Required. A member's position, or index, within a collection.|
 **Settings**
The  _collectionindex_ can be either a string or an integer. If it is a string, it must be a valid member name. If it is an integer, the minimum value is 0 and the maximum value is one less than the number of items in the collection.
 **Remarks**
If an invalid index or name is specified, an error occurs.

