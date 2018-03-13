---
title: Shape.ForeignType Property (Visio)
keywords: vis_sdr.chm11213555
f1_keywords:
- vis_sdr.chm11213555
ms.prod: visio
api_name:
- Visio.Shape.ForeignType
ms.assetid: a6cda280-bf0c-b8b0-0750-0ec5fbad90e0
ms.date: 06/08/2017
---


# Shape.ForeignType Property (Visio)

Returns the subtype of a  **Shape** object that represents a foreign object. Read-only.


## Syntax

 _expression_ . **ForeignType**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Integer


## Remarks

If the  **Type** property of a **Shape** object returns any value other than **visTypeForeignObject** , the **ForeignType** property returns the same value as the **Shape** object's **Type** property. If the **Type** property of a **Shape** object returns **visTypeForeignObject** , the **ForeignType** property returns a combination of the following values.



| **Constant **                      | **Value ** |
|:-----------------------------------|:-----------|
| <strong>visTypeMetafile</strong>   | &;H0010    |
| <strong>visTypeBitmap</strong>     | &;H0020    |
| <strong>visTypeIsLinked</strong>   | &;H0100    |
| <strong>visTypeIsEmbedded</strong> | &;H0200    |
| <strong>visTypeIsControl</strong>  | &;H0400    |
| <strong>visTypeIsOLE2</strong>     | &;H8000    |

If the shape represents an OLE 2.0 embedded object, for example, its  **ForeignType** property is &;H8200.


