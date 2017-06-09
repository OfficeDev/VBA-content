---
title: Cell.ResultForce Property (Visio)
keywords: vis_sdr.chm10114200
f1_keywords:
- vis_sdr.chm10114200
ms.prod: visio
api_name:
- Visio.Cell.ResultForce
ms.assetid: 96579953-05f2-edf5-02d6-54ef0e632215
ms.date: 06/08/2017
---


# Cell.ResultForce Property (Visio)

Sets a cell's value, even if the cell's formula is protected with the GUARD function. Read/write.


## Syntax

 _expression_ . **ResultForce**( **_UnitsNameOrCode_** )

 _expression_ A variable that represents a **Cell** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Required| **Variant**|The units to use when setting the cell's value.|

### Return Value

Double


## Remarks

Use the  **ResultForce** method to set a cell's value even if the cell's formula is protected with a GUARD function. If the string is invalid, an error is generated.

Setting the  **ResultForce** property is similar to setting a cell's **ResultFromIntForce** property. The difference is that the **ResultFromIntForce** property accepts an integer for the value of the cell, whereas the **ResultForce** property accepts a floating point number.

You can specify  _UnitsNameOrCode_ as an integer or a string value. For example, the following statements all set _UnitsNameOrCode_ to inches.

 **Cell.ResultForce** ( **visInches** ) = _newValue_

 **Cell.ResultForce** (65) = _newValue_

 **Cell.ResultForce** ("in") = _newValue_ where "in" can also be any of the alternate strings representing inches, such as "inch", "in.", or "intCounter".

For a complete list of valid unit strings along with their corresponding Automation constants (integer values), see [About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).

Automation constants for representing units are declared by the Visio type library in member  **[VisUnitCodes ](visunitcodes-enumeration-visio.md)** .

To specify internal units, pass a zero-length string (""). Internal units are inches for distance and radians for angles. To specify implicit units, you must use the  **Formula** property.


