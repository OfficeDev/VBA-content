---
title: ContainerProperties.RemoveMember Method (Visio)
keywords: vis_sdr.chm17662335
f1_keywords:
- vis_sdr.chm17662335
ms.prod: visio
api_name:
- Visio.ContainerProperties.RemoveMember
ms.assetid: 953beb58-ea8a-7c1f-20c1-0fe4de23e831
ms.date: 06/08/2017
---


# ContainerProperties.RemoveMember Method (Visio)

Removes a shape or set of shapes from the container.


## Syntax

 _expression_ . **RemoveMember**( **_ObjectToRemove_** )

 _expression_ A variable that represents a **[ContainerProperties](containerproperties-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectToRemove_|Required| **[UNKNOWN]**|The shape or shapes to remove from the container. Can be a  **[Shape](shape-object-visio.md)** or **[Selection](selection-object-visio.md)** selection.|

### Return Value

 **Nothing**


## Remarks

The  **RemoveMember** method removes from the container the shapes specified in the _ObjectToRemove_ parameter.

If the container is a list, Microsoft Visio removes the shapes specified in  _ObjectToRemove_ both from the list (if it is a list member) and from the list container.

If the  **[ContainerProperties.LockMembership](containerproperties-lockmembership-property-visio.md)** property is **True** , Visio returns a Disabled error.

If  _ObjectToRemove_ does not contain top-level shapes on the page, Visio returns an Invalid Parameter error. However, if _ObjectToRemove_ is not a container member, Visio does not return an error.


