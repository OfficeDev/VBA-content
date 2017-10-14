---
title: Document.LeftMargin Property (Visio)
keywords: vis_sdr.chm10513830
f1_keywords:
- vis_sdr.chm10513830
ms.prod: visio
api_name:
- Visio.Document.LeftMargin
ms.assetid: 9f880830-8b63-2a34-2a02-fd6b6a225c7a
ms.date: 06/08/2017
---


# Document.LeftMargin Property (Visio)

Specifies the left margin, which is used when printing. Read/write.


## Syntax

 _expression_ . **LeftMargin**( **_UnitsNameOrCode_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Optional| **Variant**|The units to use when retrieving or setting the margin value. Defaults to internal drawing units.|

### Return Value

Double


## Remarks

The  **LeftMargin** property corresponds to the **Left** setting in the **Print Setup** dialog box (on the **Design** tab, click the **Page Setup** arrow, and then, on the **Print Setup** tab, click **Setup**).

You can specify  _UnitsNameOrCode_ as an integer or a string value. If the string is invalid, an error is generated. For example, the following statements all set _UnitsNameOrCode_ to inches.

 **Document.LeftMargin** ( **visInches** ) = _newValue_

 **Document.LeftMargin** (65) = _newValue_

 **Document.LeftMargin** ("in") = _newValue_ where "in" can also be any of the alternate strings representing inches, such as "inch", "in.", or "i".

For a complete list of valid unit strings along with corresponding Automation constants (integer values), see [About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).

Automation constants for representing units are declared by the Microsoft Visio type library in member  **[VisUnitCodes](visunitcodes-enumeration-visio.md)** .


