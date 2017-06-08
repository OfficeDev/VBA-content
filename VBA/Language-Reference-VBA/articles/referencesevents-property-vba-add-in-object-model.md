---
title: ReferencesEvents Property (VBA Add-In Object Model)
keywords: vbob6.chm1092849
f1_keywords:
- vbob6.chm1092849
ms.prod: office
ms.assetid: 2482995e-ca97-067c-a7ae-cbeca2113199
ms.date: 06/08/2017
---


# ReferencesEvents Property (VBA Add-In Object Model)



Returns the  **ReferencesEvents** object. Read-only.
 **Settings**
The setting for the [argument](vbe-glossary.md) you pass to the **ReferencesEvents** property is:


|**Argument**|**Description**|
|:-----|:-----|
| _vbproject_|If  _vbproject_ points to **Nothing**, the object that is returned will supply events for the **References** collections of all **VBProject** objects in the **VBProjects** collection.If  _vbproject_ points to a valid **VBProject** object, the object that is returned will supply events for only the **References** collection for that[project](vbe-glossary.md).|
 **Remarks**
The  **ReferencesEvents** property takes an argument and returns an[event source object](vbe-glossary.md). The  **ReferencesEvents** object is the source for events that are triggered when references are added or removed.

