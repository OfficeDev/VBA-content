---
title: Cell.ResultFromIntForce Property (Visio)
keywords: vis_sdr.chm10114210
f1_keywords:
- vis_sdr.chm10114210
ms.prod: visio
api_name:
- Visio.Cell.ResultFromIntForce
ms.assetid: e22b2479-a55f-c08b-4d2b-18f8225900fa
ms.date: 06/08/2017
---


# Cell.ResultFromIntForce Property (Visio)

Sets the value of a cell to an integer value, even if the cell's formula is protected with the GUARD function. Read/write.


## Syntax

 _expression_ . **ResultFromIntForce**( **_UnitsNameOrCode_** )

 _expression_ A variable that represents a **Cell** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Required| **Variant**|The units to use when setting the cell's value.|

### Return Value

Long


## Remarks

Use the  **ResultFromIntForce** property to set a cell's value even if the cell's formula is protected with a GUARD function. Otherwise, it is identical in behavior to the **ResultFromInt** property.

Setting the  **ResultFromIntForce** property is similar to setting a cell's **ResultForce** property. The difference is that the **ResultFromIntForce** property accepts an integer for the value of the cell, whereas the **ResultForce** property accepts a floating point number.


