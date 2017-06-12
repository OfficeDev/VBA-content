---
title: Unable to activate object (Error 31027)
keywords: vblr6.chm31027
f1_keywords:
- vblr6.chm31027
ms.prod: office
ms.assetid: cfc1ae3c-83ad-a33d-2d02-3550a3ee8a95
ms.date: 06/08/2017
---


# Unable to activate object (Error 31027)

The object's source document can't be loaded, or the application that created the object isn't available.

This error occurs when you try to activate a linked object (set  **Action** = 7) and the file specified in the **SourceDoc** property has been deleted or no longer exists.

This error also occurs when you activate an object (set  **Action** = 7), and the action specified by the **Verb** property isn't valid. Some applications that provide objects may support different verbs, depending on the state of the application. All the verbs supported by an application are listed in the **ObjectVerbs** property list. However, some verbs may not be valid for the application's current state.


