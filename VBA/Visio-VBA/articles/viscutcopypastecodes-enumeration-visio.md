---
title: VisCutCopyPasteCodes Enumeration (Visio)
keywords: vis_sdr.chm70355
f1_keywords:
- vis_sdr.chm70355
ms.prod: visio
ms.assetid: 097ff931-bf8d-2d96-a718-41f7708bc265
ms.date: 06/08/2017
---


# VisCutCopyPasteCodes Enumeration (Visio)

Flags to be passed to the  **Cut** , **Copy** , or **Paste** methods.



|**Flag**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visCopyPasteNormal**|&;H0|Follow default copying behavior.|
| **visCopyPasteNoTranslate**|&;H1|Copy shapes to their original coordinate locations.|
| **visCopyPasteCenter**|&;H2|Copy shapes to the center of the page.|
| **visCopyPasteNoHealConnectors**|&;H4|Do not clean up connectors attached to cut shapes.|
| **visCopyPasteNoContainerMembers**|&;H8|Do not cut and copy unselected members of containers or lists.|
| **visCopyPasteNoAssociatedCallouts**|&;H16|Do not cut and copy unselected callouts associated with shapes.|
| **visCopyPasteDontAddToContainers**|&;H32|Do not add pasted shapes to any underlying containers.|
| **visCopyPasteNoCascade**|&;H64|Do not offset shapes on copy.|

