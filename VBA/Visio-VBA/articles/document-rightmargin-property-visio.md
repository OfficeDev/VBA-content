---
title: Document.RightMargin Property (Visio)
keywords: vis_sdr.chm10514235
f1_keywords:
- vis_sdr.chm10514235
ms.prod: visio
api_name:
- Visio.Document.RightMargin
ms.assetid: ee2fc9f4-92a6-a787-7fa0-cd43da52fadb
ms.date: 06/08/2017
---


# Document.RightMargin Property (Visio)

Specifies the right margin, which is used when printing. Read/write.


## Syntax

 _expression_ . **RightMargin**( **_UnitsNameOrCode_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Optional| **Variant**|The units to use when retrieving or setting the margin value.|

### Return Value

Double


## Remarks

If  _UnitsNameOrCode_ is not provided, the **RightMargin** property will default to internal drawing units (inches).

The  **RightMargin** property corresponds to the **Right** setting in the **Print Setup** dialog box (on the **Design** tab, click the **Page Setup** arrow, and then, on the **Print Setup** tab, click **Setup**).

You can specify  _UnitsNameOrCode_ as an integer or a string value. If the string is invalid, an error is generated. For example, the following statements all set _UnitsNameOrCode_ to inches.

 **Document.RightMargin** ( **visInches** ) = _newValue_

 **Document.RightMargin** (65) = _newValue_

 **Document.RightMargin** ("in") = _newValue_ , where "in" can also be any of the alternate strings representing inches, such as "inch", "in.", or "intCounter".

For a complete list of valid unit strings along with corresponding Automation constants (integer values), see [About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).

Automation constants for representing units are declared by the Visio type library in member  **[VisUnitCodes](visunitcodes-enumeration-visio.md)** .


