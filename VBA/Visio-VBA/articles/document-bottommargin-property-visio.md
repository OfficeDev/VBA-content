---
title: Document.BottomMargin Property (Visio)
keywords: vis_sdr.chm10513150
f1_keywords:
- vis_sdr.chm10513150
ms.prod: visio
api_name:
- Visio.Document.BottomMargin
ms.assetid: 5fd185a5-ecc9-000e-f5b0-fa309d52847a
ms.date: 06/08/2017
---


# Document.BottomMargin Property (Visio)

Specifies the bottom margin when printing the pages in a document. Read/write.


## Syntax

 _expression_ . **BottomMargin**( **_UnitsNameOrCode_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Optional| **Variant**|he units to use when retrieving or setting the margin value. Defaults to internal drawing units.|

### Return Value

Double


## Remarks

The value of this property corresponds to the value entered in the  **Bottom** box in the **Print Setup** dialog box (on the **Design** tab, click the **Page Setup** arrow, and then click **Setup** on the **Print Setup** tab).

You can specify  _UnitsNameOrCode_ as an integer or a string value. If the string is invalid, an error is generated. For example, the following statements all set _UnitsNameOrCode_ to inches.

 **ActiveDocument.BottomMargin** ( **visInches** ) = _newValue_

 **ActiveDocument.BottomMargin** (65) = _newValue_

 **ActiveDocument.BottomMargin** ("in") = _newValue_ where "in" can also be any of the alternate strings representing inches, such as "inch", "in.", or "i".

For a complete list of valid unit strings along with corresponding Automation constants (integer values), see [About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).

Automation constants for representing units are declared by the Microsoft Visio type library in member  **[VisUnitCodes](visunitcodes-enumeration-visio.md)** .


