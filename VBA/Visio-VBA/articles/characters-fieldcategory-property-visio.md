---
title: Characters.FieldCategory Property (Visio)
keywords: vis_sdr.chm10213500
f1_keywords:
- vis_sdr.chm10213500
ms.prod: visio
api_name:
- Visio.Characters.FieldCategory
ms.assetid: b9c1ecca-ae27-83b8-862d-e8677f8c4c9a
ms.date: 06/08/2017
---


# Characters.FieldCategory Property (Visio)

Returns the field category for a field represented by an object. Read-only.


## Syntax

 _expression_ . **FieldCategory**

 _expression_ A variable that represents a **Characters** object.


### Return Value

Integer


## Remarks

If the  **Characters** object does not contain a field or contains non-field characters, the **FieldCategory** property returns an exception. Check the **IsField** property of the **Characters** object before getting its **FieldCategory** property.

Field categories correspond to those in the  **Category** list in the **Field** dialog box (click **Field** on the **Insert** tab).

To add a custom field, use the  **AddCustomField** method.

The following constants for field categories are declared by the Visio type library in  **VisFieldCategories** .



|**Constant **|**Value **|
|:-----|:-----|
| **visFCatCustom**|0 |
| **visFCatDateTime**|1 |
| **visFCatDocument**|2 |
| **visFCatGeometry**|3 |
| **visFCatObject**|4 |
| **visFCatPage**|5 |

