---
title: Category.CategoryBorderColor Property (Outlook)
keywords: vbaol11.chm3266
f1_keywords:
- vbaol11.chm3266
ms.prod: outlook
api_name:
- Outlook.Category.CategoryBorderColor
ms.assetid: 95251459-f216-7cc8-55ef-c939090cf3bf
ms.date: 06/08/2017
---


# Category.CategoryBorderColor Property (Outlook)

Returns an  **OLE_COLOR** value that represents the border color of the color swatch displayed for a **[Category](category-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **CategoryBorderColor**

 _expression_ A variable that represents a **Category** object.


## Remarks

Setting the  **[Color](category-color-property-outlook.md)** property of the **Category** object to an **[OlCategoryColor](olcategorycolor-enumeration-outlook.md)** constant changes the value of this property, as well as the value of the **[CategoryGradientBottomColor](category-categorygradientbottomcolor-property-outlook.md)** and **[CategoryGradientTopColor](category-categorygradienttopcolor-property-outlook.md)** properties.

This color should be used to display the border of a gradient-shaded color swatch for the  **Category** object.


## See also


#### Concepts


[Category Object](category-object-outlook.md)

