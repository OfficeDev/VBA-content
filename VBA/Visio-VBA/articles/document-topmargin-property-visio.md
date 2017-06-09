---
title: Document.TopMargin Property (Visio)
keywords: vis_sdr.chm10514580
f1_keywords:
- vis_sdr.chm10514580
ms.prod: visio
api_name:
- Visio.Document.TopMargin
ms.assetid: ed8d16c2-f80d-d444-28a4-d9f0db4ab6d3
ms.date: 06/08/2017
---


# Document.TopMargin Property (Visio)

Specifies the top margin when printing a document. Read/write.


## Syntax

 _expression_ . **TopMargin**( **_UnitsNameOrCode_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Optional| **Variant**|The units to use when retrieving or setting the margin value.|

### Return Value

Double


## Remarks

If  _UnitsNameOrCode_ is not provided, the **TopMargin** property will default to internal drawing units (inches).

The  **TopMargin** property corresponds to the **Top** setting in the **Print Setup** dialog box (on the **Design** tab, click the **Page Setup** arrow, and then click **Setup** on the **Print Setup** tab).

Units can be an integer or string value such as "inches", "inch", "in.", or "i". Strings may be used for all supported Microsoft Visio units such as centimeters, meters, miles, and so on. You can also use any of the units constants declared by the Visio type library in  **[VisUnitCodes](visunitcodes-enumeration-visio.md)** .

For a list of valid integer and string values, see [About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).


