---
title: Document.GlueSettings Property (Visio)
keywords: vis_sdr.chm10550625
f1_keywords:
- vis_sdr.chm10550625
ms.prod: visio
api_name:
- Visio.Document.GlueSettings
ms.assetid: 192fb40f-d244-48e9-59ad-4439385bf3e5
ms.date: 06/08/2017
---


# Document.GlueSettings Property (Visio)

Determines the objects that shapes glue to when glue is enabled in the document. Read/write.


## Syntax

 _expression_ . **GlueSettings**

 _expression_ A variable that represents a **Document** object.


### Return Value

VisGlueSettings


## Remarks

Setting the value of the  **GlueSettings** property is equivalent to selecting options under **Glue to** on the **General** tab in the **Snap &; Glue** dialog box (on the **View** tab, click the **Visual Aids** arrow).

The  **GlueSettings** property can be any combination of the following **VisGlueSettings** constants, which are declared in the Microsoft Visio type library.



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visGlueToNone**|&;H0 |Glue is enabled but no other glue settings are on. |
| **visGlueToGuides**|&;H1 |Glue to guides. |
| **visGlueToHandles**|&;H2 |Glue to shape handles. |
| **visGlueToVertices**|&;H4 |Glue to shape vertices. |
| **visGlueToConnectionPoints**|&;H8 |Glue to connection points. |
| **visGlueToGeometry**|&;H20 |Glue to shape geometry. |
| **visGlueToDisabled**|&;H8000 |Disable glue. |

