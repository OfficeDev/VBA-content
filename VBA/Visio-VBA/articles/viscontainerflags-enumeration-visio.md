---
title: VisContainerFlags Enumeration (Visio)
keywords: vis_sdr.chm70620
f1_keywords:
- vis_sdr.chm70620
ms.prod: visio
api_name:
- Visio.VisContainerFlags
ms.assetid: c440c15a-5dd9-7ece-9175-dd92283455a4
ms.date: 06/08/2017
---


# VisContainerFlags Enumeration (Visio)

Specifies which container member shape IDs to return; constants passed to the  **[ContainerProperties.GetMemberShapes](containerproperties-getmembershapes-method-visio.md)** method.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visContainerFlagsDefault**|0|Returns all shape types and includes items in nested containers.|
| **visContainerFlagsExcludeContainers**|1|Excludes member shapes that are containers.|
| **visContainerFlagsExcludeConnectors**|2|Excludes member shapes that are connectors.|
| **visContainerFlagsExcludeCallouts**|4|Excludes member shapes that are callouts.|
| **visContainerFlagsExcludeElements**|8|Excludes member shapes that are not containers, connectors, or callouts.|
| **visContainerFlagsExcludeNested**|16|Excludes any member shapes that are members of containers nested within the container.|
| **visContainerFlagsExcludeListMembers**|32|Excludes members of a list container that are explicitly members of the list. Does not exclude other shapes in the list container.|

