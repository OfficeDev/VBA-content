---
title: ContainerProperties.LockMembership Property (Visio)
keywords: vis_sdr.chm17662605
f1_keywords:
- vis_sdr.chm17662605
ms.prod: visio
api_name:
- Visio.ContainerProperties.LockMembership
ms.assetid: b82455fc-f3cb-66de-c022-ac6f63f5b4b2
ms.date: 06/08/2017
---


# ContainerProperties.LockMembership Property (Visio)

Gets or sets a value that determines whether container members can be added, removed, or deleted. Read/write.


## Syntax

 _expression_ . **LockMembership**

 _expression_ An expression that returns a **[ContainerProperties](containerproperties-object-visio.md)** object.


### Return Value

 **Boolean**


## Remarks

For normal (non-list) containers, setting  **LockMembership** to **True** does not prevent moving a container member; the container expands to include the moved member, assuming that the setting of **[ResizeAsNeeded](containerproperties-resizeasneeded-property-visio.md)** is **visContainerAutoResizeExpand** or **visContainerAutoResizeExpandContract** . If **[ResizeAsNeeded](containerproperties-resizeasneeded-property-visio.md)** is **visContainerAutoResizeNone** and you move a container member outside the outlines of the container, the shape remains a member, but the container does not expand to visibly contain it. You cannot add, remove, or delete a member from a locked container.

For list containers, setting  **LockMembership** to **True** locks container members that are also list members, which prevents moving them and thus reordering of the members. It does not, however, prevent you from moving normal (non-list) members of the container out of the list (although not out of the container). You can also delete normal members.

The setting of the  **LockMembership** property corresponds to the setting of **Lock Container** in the **Membership** group on the **Container Tools Format** tab.


