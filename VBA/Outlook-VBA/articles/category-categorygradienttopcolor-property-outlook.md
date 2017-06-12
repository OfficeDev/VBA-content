---
title: Category.CategoryGradientTopColor Property (Outlook)
keywords: vbaol11.chm3267
f1_keywords:
- vbaol11.chm3267
ms.prod: outlook
api_name:
- Outlook.Category.CategoryGradientTopColor
ms.assetid: deb7a986-8afd-465c-ed8e-3cf669f96a35
ms.date: 06/08/2017
---


# Category.CategoryGradientTopColor Property (Outlook)

Returns an  **OLE_COLOR** value that represents the top gradient color of the color swatch displayed for a **[Category](category-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **CategoryGradientTopColor**

 _expression_ A variable that represents a **Category** object.


## Remarks

Setting the  **[Color](category-color-property-outlook.md)** property of the **Category** object to an **[OlCategoryColor](olcategorycolor-enumeration-outlook.md)** constant changes the value of this property, as well as the value of the **[CategoryGradientBottomColor](category-categorygradientbottomcolor-property-outlook.md)** and **[CategoryBorderColor](category-categorybordercolor-property-outlook.md)** properties.

This color should be used to display a gradient-shaded color swatch for the  **Category** object.


## See also


#### Concepts


[Category Object](category-object-outlook.md)

