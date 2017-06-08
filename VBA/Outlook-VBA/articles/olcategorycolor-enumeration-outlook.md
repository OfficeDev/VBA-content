---
title: OlCategoryColor Enumeration (Outlook)
keywords: vbaol11.chm3119
f1_keywords:
- vbaol11.chm3119
ms.prod: outlook
api_name:
- Outlook.OlCategoryColor
ms.assetid: 048bbc6b-c49f-68a3-ac59-b61204e5ef78
ms.date: 06/08/2017
---


# OlCategoryColor Enumeration (Outlook)

Indicates the color that is specified for a category or a font in a view.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olCategoryColorBlack**|15|Black|
| **olCategoryColorBlue**|8|Blue|
| **olCategoryColorDarkBlue**|23|Dark Blue|
| **olCategoryColorDarkGray**|14|Dark Gray|
| **olCategoryColorDarkGreen**|20|Dark Green|
| **olCategoryColorDarkMaroon**|25|Dark Maroon|
| **olCategoryColorDarkOlive**|22|Dark Olive|
| **olCategoryColorDarkOrange**|17|Dark Orange|
| **olCategoryColorDarkPeach**|18|Dark Peach|
| **olCategoryColorDarkPurple**|24|Dark Purple|
| **olCategoryColorDarkRed**|16|Dark Red|
| **olCategoryColorDarkSteel**|12|Dark Steel|
| **olCategoryColorDarkTeal**|21|Dark Teal|
| **olCategoryColorDarkYellow**|19|Dark Yellow|
| **olCategoryColorGray**|13|Gray|
| **olCategoryColorGreen**|5|Green|
| **olCategoryColorMaroon**|10|Maroon|
| **olCategoryColorNone**|0|No color assigned.|
| **olCategoryColorOlive**|7|Olive|
| **olCategoryColorOrange**|2|Orange|
| **olCategoryColorPeach**|3|Peach|
| **olCategoryColorPurple**|9|Purple|
| **olCategoryColorRed**|1|Red|
| **olCategoryColorSteel**|11|Steel|
| **olCategoryColorTeal**|6|Teal|
| **olCategoryColorYellow**|4|Yellow|

## Remarks

Used by the [Color](category-color-property-outlook.md) property of the[Category Object (Outlook)](category-object-outlook.md), and the [ExtendedColor](viewfont-extendedcolor-property-outlook.md) property of the[ViewFont Object (Outlook)](viewfont-object-outlook.md).

The color constants provided here are approximations of the actual colors used by the  **Category** object. Use the[CategoryBorderColor](category-categorybordercolor-property-outlook.md), [CategoryGradientBottomColor](category-categorygradientbottomcolor-property-outlook.md), and [CategoryGradientTopColor](category-categorygradienttopcolor-property-outlook.md) properties to retrieve the **OLE_COLOR** color values that are used to represent the **Category** object, after setting the **Color** property to the appropriate constant.


