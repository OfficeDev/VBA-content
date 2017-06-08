---
title: GraphicItems.Stat Property (Visio)
keywords: vis_sdr.chm16814420
f1_keywords:
- vis_sdr.chm16814420
ms.prod: visio
api_name:
- Visio.GraphicItems.Stat
ms.assetid: 5aed1205-421f-5dbd-8cc3-dd08dcbc5e4d
ms.date: 06/08/2017
---


# GraphicItems.Stat Property (Visio)

Returns status information for an object. Read-only.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **Stat**

 _expression_ A variable that represents a **GraphicItems** object.


### Return Value

Integer


## Remarks

If an object is a reference to an entity in a document, and if that document closes, the  **Stat** property returns a value in which the **visStatClosed** bit is set.

If an object is a reference to an entity that has been deleted, the  **Stat** property returns a value in which the **visStatDeleted** bit is set.

A Component Object Model (COM) object, such as a Microsoft Visio  **Document** object, lives as long as it is held (pointed to) by a client, even if the object is logically in a deleted or closed state.


