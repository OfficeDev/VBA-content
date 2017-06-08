---
title: Characters.FieldFormat Property (Visio)
keywords: vis_sdr.chm10213510
f1_keywords:
- vis_sdr.chm10213510
ms.prod: visio
api_name:
- Visio.Characters.FieldFormat
ms.assetid: 298ee3a7-a81e-c10d-e978-ce28ca9408be
ms.date: 06/08/2017
---


# Characters.FieldFormat Property (Visio)

Returns the field format for a field represented by an object. Read-only.


## Syntax

 _expression_ . **FieldFormat**

 _expression_ A variable that represents a **Characters** object.


### Return Value

Integer


## Remarks

If the  **Characters** object does not contain a field or contains non-field characters, the **FieldFormat** property returns an exception. Check the **IsField** property of the **Characters** object before getting its **FieldFormat** property.

Field formats correspond to the formats in the  **Format** list in the **Field** dialog box (click **Field** on the **Insert** tab).

Constants for field formats are declared by the Microsoft Visio type library in  **[VisFieldFormats](visfieldformats-enumeration-visio.md)** .


