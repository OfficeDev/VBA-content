---
title: Document.HeaderMargin Property (Visio)
keywords: vis_sdr.chm10550650
f1_keywords:
- vis_sdr.chm10550650
ms.prod: visio
api_name:
- Visio.Document.HeaderMargin
ms.assetid: 7d2c137d-6b75-9747-5a6a-5e5d99156d45
ms.date: 06/08/2017
---


# Document.HeaderMargin Property (Visio)

Gets or sets the margin of a document's header. Read/write.


## Syntax

 _expression_ . **HeaderMargin**( **_UnitsNameOrCode_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Optional| **Variant**|The units to use when retrieving or setting the cell's value. Defaults to internal drawing units (inches).|

### Return Value

Double


## Remarks

You can also set this value in the  **Margin** box under **Header** in the **Header and Footer** dialog box (click the **File** tab, click **Print**, click  **Print Preview**, and then in the  **Preview** group, click **Header &; Footer**).

Automation constants for representing units are declared by the Visio type library in member  **[VisUnitCodes](visunitcodes-enumeration-visio.md)** .

For a complete list of valid unit strings along with corresponding Automation constants (integer values), see [About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).


