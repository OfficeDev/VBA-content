---
title: VisMemberAddOptions Enumeration (Visio)
keywords: vis_sdr.chm70615
f1_keywords:
- vis_sdr.chm70615
ms.prod: visio
api_name:
- Visio.VisMemberAddOptions
ms.assetid: e6833a87-2d08-19a4-c2f9-86803ca4e4ba
ms.date: 06/08/2017
---


# VisMemberAddOptions Enumeration (Visio)

Specifies whether to expand the container to accomodate the new member(s) or to resize it automatically according to the default settings; constants passed to the  **[ContainerProperties.AddMember](containerproperties-addmember-method-visio.md)** method.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visMemberAddUseResizeSetting**|0|Defer to the setting of the  **[ContainerProperties.ResizeAsNeeded](containerproperties-resizeasneeded-property-visio.md)** property.|
| **visMemberAddExpandContainer**|1|Expand the container to fit the incoming shape(s).|
| **visMemberAddDoNotExpand**|2|Do not expand the container to fit the incoming shape(s).|

