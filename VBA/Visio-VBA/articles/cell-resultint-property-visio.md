---
title: Cell.ResultInt Property (Visio)
keywords: vis_sdr.chm10114215
f1_keywords:
- vis_sdr.chm10114215
ms.prod: visio
api_name:
- Visio.Cell.ResultInt
ms.assetid: f3e2ef7d-cde1-a0d4-3d02-f5bf329cd0c3
ms.date: 06/08/2017
---


# Cell.ResultInt Property (Visio)

Gets the value of a cell expressed as an integer. Read-only.


## Syntax

 _expression_ . **ResultInt**( **_UnitsNameOrCode_** , **_fRound_** )

 _expression_ A variable that represents a **Cell** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Required| **Variant**|The units to use when retrieving the cell's value.|
| _fRound_|Required| **Integer**|Zero (0) to truncate the value; non-zero to round it.|

### Return Value

Long


## Remarks

Getting the  **ResultInt** property is similar to a getting a cell's **Result** property. The difference is that the **ResultInt** property returns an integer for the value of the cell, whereas the **Result** property returns a floating point number.

You can specify  _UnitsNameOrCode_ as an integer or a string value. If the string is invalid, an error is generated. For example, the following statements all set _UnitsNameOrCode_ to inches.

 _lngRet_ = **Cell.ResultInt** ( **visInches** , _fRound_)

 _lngRet_ = **Cell.ResultInt** (65, _fRound_)

 _lngRet_ = **Cell.ResultInt** ("in", _fRound_) where "in" can also be any of the alternate strings representing inches, such as "inch", "in.", or "intCounter".

For a complete list of valid unit strings along with their corresponding Automation constants (integer values), see [About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).

Automation constants for representing units are declared by the Visio type library in member  **[VisUnitCodes ](visunitcodes-enumeration-visio.md)** .

The following constants for  _fRound_ are declared in the Visio type library in member **VisRoundFlags** .



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visTruncate**|0 |Truncate the result. |
| **visRound**|1 |Round the result. |

