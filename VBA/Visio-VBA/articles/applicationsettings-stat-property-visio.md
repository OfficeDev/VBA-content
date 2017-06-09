---
title: ApplicationSettings.Stat Property (Visio)
keywords: vis_sdr.chm16214420
f1_keywords:
- vis_sdr.chm16214420
ms.prod: visio
api_name:
- Visio.ApplicationSettings.Stat
ms.assetid: dd322ca5-6f48-94ab-8632-f60896dd3228
ms.date: 06/08/2017
---


# ApplicationSettings.Stat Property (Visio)

Returns status information for an object. Read-only.


## Syntax

 _expression_ . **Stat**

 _expression_ A variable that represents an **ApplicationSettings** object.


### Return Value

Integer


## Remarks

Possible values returned by the  **Stat** property are declared by the Visio type library in **[VisStatCodes](visstatcodes-enumeration-visio.md)** .

If an object is a reference to an entity in a document, and if that document closes, the  **Stat** property returns a value in which the **visStatClosed** bit is set.

If an object is a reference to an entity that has been deleted, the  **Stat** property returns a value in which the **visStatDeleted** bit is set.

A Component Object Model (COM) object, such as a Microsoft Visio  **Document** object, lives as long as it is held (pointed to) by a client, even if the object is logically in a deleted or closed state.


