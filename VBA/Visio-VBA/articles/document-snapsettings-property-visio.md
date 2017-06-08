---
title: Document.SnapSettings Property (Visio)
keywords: vis_sdr.chm10550890
f1_keywords:
- vis_sdr.chm10550890
ms.prod: visio
api_name:
- Visio.Document.SnapSettings
ms.assetid: c3ced586-d9c7-01bd-6b32-99fedda3c2b8
ms.date: 06/08/2017
---


# Document.SnapSettings Property (Visio)

Determines the objects that shapes snap to when snap is active in the document. Read/write.


## Syntax

 _expression_ . **SnapSettings**

 _expression_ A variable that represents a **Document** object.


### Return Value

VisSnapSettings


## Remarks

The value of the  **SnapSettings** property is equivalent to selecting check boxes under **Snap to** on the **General** tab of the **Snap &; Glue** dialog box (click the **Visual Aids** arrow on the **View** tab).

The  **SnapSettings** property can be any combination of the following **VisSnapSettings** constants, which are declared in the Visio type library.



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visSnapToNone**|&;H0 |Snap to nothing. |
| **visSnapToRulerSubdivisions**|&;H1 |Snap to tick marks on the ruler. |
| **visSnapToGrid**|&;H2 |Snap to the grid. |
| **visSnapToGuides**|&;H4 |Snap to guides. |
| **visSnapToHandles**|&;H8 |Snap to selection handles. |
| **visSnapToVertices**|&;H10 |Snap to vertices. |
| **visSnapToConnectionPoints**|&;H20 |Snap to connection points. |
| **visSnapToGeometry**|&;H100 |Snap to the visible edges of shapes. |
| **visSnapToAlignmentBox**|&;H200 |Snap to the alignment box. |
| **visSnapToExtensions**|&;H400 |Snap to shape extensions options. |
| **visSnapToDisabled**|&;H8000 |Disable snap. |
| **visSnapToIntersections**|&;H10000 |Snap to intersections. |

