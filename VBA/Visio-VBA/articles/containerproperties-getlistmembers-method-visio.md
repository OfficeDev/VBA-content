---
title: ContainerProperties.GetListMembers Method (Visio)
keywords: vis_sdr.chm17662345
f1_keywords:
- vis_sdr.chm17662345
ms.prod: visio
api_name:
- Visio.ContainerProperties.GetListMembers
ms.assetid: 9aa6047a-ae20-d05c-cb59-56594ed08b2f
ms.date: 06/08/2017
---


# ContainerProperties.GetListMembers Method (Visio)

Returns an array of shape identifiers (IDs) of member shapes in the list.


## Syntax

 _expression_ . **GetListMembers**

 _expression_ A variable that represents a **[ContainerProperties](containerproperties-object-visio.md)** object.


### Return Value

 **Long()**


## Remarks

 **GetListMembers** returns an empty array if there are no shapes in the list.

If the container is not a list, Microsoft Visio returns an Invalid Source error.


