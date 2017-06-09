---
title: Remove Method (VBA Add-In Object Model)
keywords: vbob6.chm100142
f1_keywords:
- vbob6.chm100142
ms.prod: office
ms.assetid: acc163b9-e5ad-ef39-013a-614fc24bcde1
ms.date: 06/08/2017
---


# Remove Method (VBA Add-In Object Model)



Removes an item from a [collection](vbe-glossary.md).
 **Syntax**
 _object_**.Remove(**_component_**)**
The  **Remove** syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. An [object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.|
| _component_|Required. For the  **LinkedWindows** collection, an object. For the **References** collection, a reference to a[type library](vbe-glossary.md) or a[project](vbe-glossary.md). For the  **VBComponents** collection, an enumerated[constant](vbe-glossary.md) representing a[class module](vbe-glossary.md), a form, or a [standard module](vbe-glossary.md). For the  **VBProjects** collection, a standalone project.|
 **Remarks**
When used on the  **LinkedWindows** collection, the **Remove** method removes a window from the collection of currently[linked windows](vbe-glossary.md). The removed window becomes a floating window that has its own [linked window frame](vbe-glossary.md). The  **Remove** method can only be used on a standalone project. It generates a run-time error if you try to use it on a host project.


 **Important**  Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements generate run-time errors when run on the Macintosh.



