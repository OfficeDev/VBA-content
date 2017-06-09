---
title: Category.CategoryGradientBottomColor Property (Outlook)
keywords: vbaol11.chm3268
f1_keywords:
- vbaol11.chm3268
ms.prod: outlook
api_name:
- Outlook.Category.CategoryGradientBottomColor
ms.assetid: 5f082300-2eb0-b297-dc54-9657da5ae319
ms.date: 06/08/2017
---


# Category.CategoryGradientBottomColor Property (Outlook)

Returns an  **OLE_COLOR** value that represents the bottom gradient color of the color swatch displayed for a **[Category](category-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **CategoryGradientBottomColor**

 _expression_ A variable that represents a **Category** object.


## Remarks

Setting the  **[Color](category-color-property-outlook.md)** property of the **Category** object to an **[OlCategoryColor](olcategorycolor-enumeration-outlook.md)** constant changes the value of this property, as well as the value of the **[CategoryGradientTopColor](category-categorygradienttopcolor-property-outlook.md)** and **[CategoryBorderColor](category-categorybordercolor-property-outlook.md)** properties.

This color should be used to display a gradient-shaded color swatch for the  **Category** object.


## See also


#### Concepts


[Category Object](category-object-outlook.md)

