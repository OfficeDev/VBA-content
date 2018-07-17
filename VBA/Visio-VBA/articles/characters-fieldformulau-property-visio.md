---
title: Characters.FieldFormulaU Property (Visio)
keywords: vis_sdr.chm10251965
f1_keywords:
- vis_sdr.chm10251965
ms.prod: visio
api_name:
- Visio.Characters.FieldFormulaU
ms.assetid: 83a6f079-bd1a-7512-61f1-0b9fa7c83964
ms.date: 06/08/2017
---


# Characters.FieldFormulaU Property (Visio)

Returns the universal-syntax formula of the custom field represented by an object. Read-only.


## Syntax

 _expression_ . **FieldFormulaU**

 _expression_ A variable that represents a **Characters** object.


### Return Value

String


## Remarks

If the  **Characters** object does not contain a field or contains non-field characters, or if the field is not a custom field, the **FieldFormulaU** property returns an exception. Check the **IsField** and **FieldCategory** properties of the **Characters** object before getting its **FieldFormulaU** property.

The formula returned by the  **FieldFormulaU** property corresponds to the formula that appears in the **Custom formula** box in the **Field** dialog box (click **Field** on the **Insert** tab).




 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **FieldFormula** property to get a formula in local syntax. Use the **FieldFormulaU** property to get a formula in universal syntax.


