---
title: IndexedValue Property (VBA Add-In Object Model)
keywords: vbob6.chm1099626
f1_keywords:
- vbob6.chm1099626
ms.prod: office
ms.assetid: df4356f9-aee9-ead5-82ef-185e4128c4fc
ms.date: 06/08/2017
---


# IndexedValue Property (VBA Add-In Object Model)



Returns or sets a value for a member of a [property](vbe-glossary.md) that is an indexed list or an[array](vbe-glossary.md).
 **Remarks**
The value returned or set by the  **IndexedValue** property is an[expression](vbe-glossary.md) that evaluates to a type that is accepted by the object. For a property that is an indexed list or array, you must use the **IndexedValue** property instead of the **Value** property. An indexed list is a[numeric expression](vbe-glossary.md) specifying index position.
 **IndexedValue** accepts up to 4 indices. The number of indices accepted by **IndexedValue** is the value returned by the **NumIndices** property.
The  **IndexedValue** property is used only if the value of the **NumIndices** property is greater than zero. Values in indexed lists are set or returned with a single index.

