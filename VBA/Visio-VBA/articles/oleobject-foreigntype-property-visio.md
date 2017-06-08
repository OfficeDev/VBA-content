---
title: OLEObject.ForeignType Property (Visio)
keywords: vis_sdr.chm15213555
f1_keywords:
- vis_sdr.chm15213555
ms.prod: visio
api_name:
- Visio.OLEObject.ForeignType
ms.assetid: efbbf903-12ba-e269-bb86-eb4ecc99e190
ms.date: 06/08/2017
---


# OLEObject.ForeignType Property (Visio)

Returns the subtype of a  **Shape** object that represents a foreign object. Read-only.


## Syntax

 _expression_ . **ForeignType**

 _expression_ A variable that represents an **OLEObject** object.


### Return Value

Integer


## Remarks

If the  **Type** property of a **Shape** object returns any value other than **visTypeForeignObject** , the **ForeignType** property returns the same value as the **Shape** object's **Type** property. If the **Type** property of a **Shape** object returns **visTypeForeignObject** , the **ForeignType** property returns a combination of the following values.



|**Constant **|**Value **|
|:-----|:-----|
| **visTypeMetafile**|&;H0010|
| **visTypeBitmap**|&;H0020|
| **visTypeIsLinked**|&;H0100|
| **visTypeIsEmbedded**|&;H0200|
| **visTypeIsControl**|&;H0400|
| **visTypeIsOLE2**|&;H8000|
If the shape represents an OLE 2.0 embedded object, for example, its  **ForeignType** property is &;H8200.


