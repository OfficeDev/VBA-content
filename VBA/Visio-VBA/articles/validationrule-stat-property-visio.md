---
title: ValidationRule.Stat Property (Visio)
keywords: vis_sdr.chm18414420
f1_keywords:
- vis_sdr.chm18414420
ms.prod: visio
api_name:
- Visio.ValidationRule.Stat
ms.assetid: efcf067c-8adf-8983-6521-7fbe76a289ca
ms.date: 06/08/2017
---


# ValidationRule.Stat Property (Visio)

Returns status information for an object. Read-only.


## Syntax

 _expression_ . **Stat**

 _expression_ A variable that represents a **[ValidationRule](validationrule-object-visio.md)** object.


### Return Value

 **[VisStatCodes](visstatcodes-enumeration-visio.md)**


## Remarks

If an object is a reference to an entity in a document, and if that document closes, the  **Stat** property returns a value in which the **visStatClosed** bit is set.

If an object is a reference to an entity that has been deleted, the  **Stat** property returns a value in which the **visStatDeleted** bit is set.

A Component Object Model (COM) object, such as a Microsoft Visio  **[Document](document-object-visio.md)** object, lives as long as it is held (pointed to) by a client, even if the object is logically in a deleted or closed state.


