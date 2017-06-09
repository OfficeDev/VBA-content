---
title: Characters.FieldCode Property (Visio)
keywords: vis_sdr.chm10213505
f1_keywords:
- vis_sdr.chm10213505
ms.prod: visio
api_name:
- Visio.Characters.FieldCode
ms.assetid: 901e6617-2e4b-6f99-f886-e3c7348a306d
ms.date: 06/08/2017
---


# Characters.FieldCode Property (Visio)

Returns the field code for a field represented by an object. Read-only.


## Syntax

 _expression_ . **FieldCode**

 _expression_ A variable that represents a **Characters** object.


### Return Value

Integer


## Remarks

If the  **Characters** object does not contain a field or contains non-field characters, the **FieldCode** property returns an exception. Check the **IsField** property of the **Characters** object before getting its **FieldCode** property.

Field codes correspond to the fields in the  **Field** list in the **Field** dialog box (click **Field** on the **Insert** tab).

Constants for field codes are declared by the Microsoft Visio type library in  **[VisFieldCodes](visfieldcodes-enumeration-visio.md)** .


