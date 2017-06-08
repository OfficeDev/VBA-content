---
title: Document.PaperWidth Property (Visio)
keywords: vis_sdr.chm10514025
f1_keywords:
- vis_sdr.chm10514025
ms.prod: visio
api_name:
- Visio.Document.PaperWidth
ms.assetid: e43d7d44-31ad-24e3-79e4-6005cbd65612
ms.date: 06/08/2017
---


# Document.PaperWidth Property (Visio)

Returns the width of a document's printed page. Read-only.


## Syntax

 _expression_ . **PaperWidth**( **_UnitsNameOrCode_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Required| **Variant**|The units to use when setting or retrieving the paper width.|

### Return Value

Double


## Remarks

The  **PaperWidth** property value can be a string such as "inches", "inch", "in.", or "i". Strings may be used for all supported Microsoft Visio units such as centimeters, meters, miles, and so on. You can also use any of the units constants declared by the Visio type library in member **[VisUnitCodes](visunitcodes-enumeration-visio.md)** .


