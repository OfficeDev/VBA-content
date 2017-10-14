---
title: AccelTable.SetID Property (Visio)
keywords: vis_sdr.chm14714315
f1_keywords:
- vis_sdr.chm14714315
ms.prod: visio
api_name:
- Visio.AccelTable.SetID
ms.assetid: d73787cc-0145-845e-6675-906d4d2aaa78
ms.date: 06/08/2017
---


# AccelTable.SetID Property (Visio)

Returns the set ID of an  **AccelTable** object in its collection. Read-only.


## Syntax

 _expression_ . **SetID**

 _expression_ A variable that represents an **AccelTable** object.


### Return Value

Long


## Remarks

Each  **AccelTable** object has a set ID that corresponds to a Microsoft Visio window context.

You can retrieve an object from its collection by passing the object's set ID to the  **ItemAtID** property. You can also set the set ID of an object by using the **AddAtID** method.

Valid set ID values are declared by the Visio type library in  **[VisUIObjSets](visuiobjsets-enumeration-visio.md)** .


