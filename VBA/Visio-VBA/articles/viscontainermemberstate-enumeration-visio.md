---
title: VisContainerMemberState Enumeration (Visio)
keywords: vis_sdr.chm70605
f1_keywords:
- vis_sdr.chm70605
ms.prod: visio
api_name:
- Visio.VisContainerMemberState
ms.assetid: 41b5c521-79b7-d7ce-38b3-17841815d429
ms.date: 06/08/2017
---


# VisContainerMemberState Enumeration (Visio)

Specifies constants that denote the state of the input member shape with respect to the container; returned by the  **[ContainerProperties.GetMemberState](containerproperties-getmemberstate-method-visio.md)** method.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visContainerMemberNotAMember**|0|The shape is not a member of the container.|
| **visContainerMemberInterior**|1|The member shape is within the bounds of the container.|
| **visContainerMemberOnBoundary**|2|The member shape is on the boundary of the container.|
| **visContainerMemberOutside**|3|The member shape is outside the bounds of the container.|
| **visContainerMemberInList**|4|The member shape is a list member.|

