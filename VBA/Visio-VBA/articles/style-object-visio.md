---
title: Style Object (Visio)
keywords: vis_sdr.chm10255
f1_keywords:
- vis_sdr.chm10255
ms.prod: visio
api_name:
- Visio.Style
ms.assetid: fdebb8d1-8910-3df8-74cd-9f847efb7ecb
ms.date: 06/08/2017
---


# Style Object (Visio)

Represents a style defined in a document.


## Remarks

You retrieve a particular style from the  **Styles** collection of a **Document** object.

The default property of a  **Style** object is **Name** .

Any  **Shape** object to which a style is applied inherits the attributes defined by the style. Use the **LineStyle** , **FillStyle** , **TextStyle** , or **Style** property of a **Shape** object to apply a style to a shape or to determine what style is applied to a shape.

Like a  **Shape** object, a **Style** object has cells whose formulas define the values of the style's attributes. To retrieve one of these cells, use the **Cells** or **CellsSRC** property of the **Style** object.


