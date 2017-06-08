---
title: Cell.FormulaForceU Property (Visio)
keywords: vis_sdr.chm10151970
f1_keywords:
- vis_sdr.chm10151970
ms.prod: visio
api_name:
- Visio.Cell.FormulaForceU
ms.assetid: 386003e3-b9e9-4c35-ac14-55bdb8da4375
ms.date: 06/08/2017
---


# Cell.FormulaForceU Property (Visio)

Sets the universal syntax formula in a  **Cell** object, even if the formula is protected with a GUARD function. Read/write.


## Syntax

 _expression_ . **FormulaForceU**

 _expression_ A variable that represents a **Cell** object.


### Return Value

String


## Remarks

Many of the SmartShapes symbols provided with Microsoft Visio have guarded cells to maintain their smart behavior. When you change the formula in a guarded cell, the shape's behavior might change in unexpected ways.




 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **FormulaForce** property when you want to use local syntax in the formula. Use the **FormulaForceU** property when you want to use universal syntax in the formula.


