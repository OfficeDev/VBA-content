---
title: Reference Object (VBA Add-In Object Model)
keywords: vbob6.chm104053
f1_keywords:
- vbob6.chm104053
ms.prod: office
ms.assetid: 559d4da0-624f-a574-575d-768155c89c72
ms.date: 06/08/2017
---


# Reference Object (VBA Add-In Object Model)



Represents a reference to a [type library](vbe-glossary.md) or a[project](vbe-glossary.md).
 **Remarks**
Use the  **Reference** object to verify whether a reference is still valid.
The  **IsBroken** property returns **True** if the reference no longer points to a valid reference. The **BuiltIn** property returns **True** if the reference is a default reference that can't be moved or removed. Use the **Name** property to determine if the reference you want to add or remove is the correct one.

