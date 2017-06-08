---
title: Document.PaperHeight Property (Visio)
keywords: vis_sdr.chm10514015
f1_keywords:
- vis_sdr.chm10514015
ms.prod: visio
api_name:
- Visio.Document.PaperHeight
ms.assetid: 305356e8-69d6-bae3-5136-d931fcf967b5
ms.date: 06/08/2017
---


# Document.PaperHeight Property (Visio)

Returns the height of a document's printed page. Read-only.


## Syntax

 _expression_ . **PaperHeight**( **_UnitsNameOrCode_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Required| **Variant**|The units to use when retrieving the paper height.|

### Return Value

Double


## Remarks

The  **PaperHeight** property value can be a string such as "inches", "inch", "in.", or "i". Strings may be used for all supported Microsoft Visio units such as centimeters, meters, miles, and so on. You can also use any of the unit constants declared by the Visio type library in **[VisUnitCodes](visunitcodes-enumeration-visio.md)** .


