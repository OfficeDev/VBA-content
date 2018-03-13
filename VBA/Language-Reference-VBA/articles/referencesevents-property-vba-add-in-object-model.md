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


| <strong>Argument</strong> | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                                                                                                         |
|:--------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>vbproject</em>        | If  <em>vbproject</em> points to <strong>Nothing</strong>, the object that is returned will supply events for the <strong>References</strong> collections of all <strong>VBProject</strong> objects in the <strong>VBProjects</strong> collection.If  <em>vbproject</em> points to a valid <strong>VBProject</strong> object, the object that is returned will supply events for only the <strong>References</strong> collection for that[project](vbe-glossary.md). |

 **Remarks**
The  **ReferencesEvents** property takes an argument and returns an[event source object](vbe-glossary.md). The  **ReferencesEvents** object is the source for events that are triggered when references are added or removed.

