---
title: Value Property (VBA Add-In Object Model)
keywords: vbob6.chm102046
f1_keywords:
- vbob6.chm102046
ms.prod: MULTIPLEPRODUCTS
ms.assetid: 9c756162-7082-7ed3-8094-6c9f24940a65
---


# Value Property (VBA Add-In Object Model)



Returns or sets a [Variant](vbe-glossary.md) specifying the value of the[property](vbe-glossary.md). Read/write.
 **Remarks**
Because the  **Value** property returns a **Variant**, you can access any property. To access a list, use the **IndexedValue** property.
If the property that the  **Property** object represents is read/write, the **Value** property is read/write. If the property is read-only, attempting to set the **Value** property causes an error. If the property is write-only, attempting to return the **Value** property causes an error.
The  **Value** property is the default property for the **Property** object.

