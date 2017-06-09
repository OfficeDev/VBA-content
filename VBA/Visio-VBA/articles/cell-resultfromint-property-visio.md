---
title: Cell.ResultFromInt Property (Visio)
keywords: vis_sdr.chm10114205
f1_keywords:
- vis_sdr.chm10114205
ms.prod: visio
api_name:
- Visio.Cell.ResultFromInt
ms.assetid: 1fb4b39b-b868-64b1-1952-405045a11d6f
ms.date: 06/08/2017
---


# Cell.ResultFromInt Property (Visio)

Sets the value of a cell to an integer value. Read/write.


## Syntax

 _expression_ . **ResultFromInt**( **_UnitsNameOrCode_** )

 _expression_ A variable that represents a **Cell** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Required| **Variant**|The units to use when setting the cell's value.|

### Return Value

Long


## Remarks

Setting the  **ResultFromInt** property is similar to setting a cell's **Result** property. The difference is that the **ResultFromInt** property accepts an integer for the value of the cell, whereas the **Result** property accepts a floating point number.

You can specify  _UnitsNameOrCode_ as an integer or a string value. If the string is invalid, an error is generated. For example, the following statements all set _UnitsNameOrCode_ to inches.

 **Cell.ResultFromInt** ( **visInches** ) = _newValue_

 **Cell.ResultFromInt** (65) = _newValue_

 **Cell.ResultFromInt** ("in") = _newValue_ where "in" can also be any of the alternate strings representing inches, such as "inch", "in.", or "intCounter".

For a complete list of valid unit strings along with their corresponding Automation constants (integer values), see [About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).

Automation constants for representing units are declared by the Visio type library in member  **[VisUnitCodes ](visunitcodes-enumeration-visio.md)** .

If the cell's formula is protected with a GUARD function, use the  **ResultFromIntForce** property.


