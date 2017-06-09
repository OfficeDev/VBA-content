---
title: ValidationRule.NameU Property (Visio)
keywords: vis_sdr.chm18451990
f1_keywords:
- vis_sdr.chm18451990
ms.prod: visio
api_name:
- Visio.ValidationRule.NameU
ms.assetid: ed5f5ba3-5dcc-2d60-45d8-8198292f2aed
ms.date: 06/08/2017
---


# ValidationRule.NameU Property (Visio)

Specifies the universal name of the  **[ValidationRule](validationrule-object-visio.md)** object. This is the default property of the object. Read/write.


## Syntax

 _expression_ . **NameU**

 _expression_ A variable that represents a **ValidationRule** object.


### Return Value

 **String**


## Remarks

You cannot assign the  **NameU** property a name that already exists in the rule set. If you attempt to do so, Visio returns an "invalid parameter" error.

You cannot set the  **NameU** property to a value that exceeds 255 characters or to an empty string.




 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to various Visio objects. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you do not want to change a name each time a solution is localized. Use the  **[Name](validationruleset-name-property-visio.md)** property to get or set an object's local name. Use the **NameU** property to get or set its universal name.


